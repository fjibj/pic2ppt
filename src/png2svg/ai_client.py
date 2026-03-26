#!/usr/bin/env python3
"""
AI Image to SVG Converter
支持多种 AI 提供商：Claude、Kimi、GLM、OpenAI
"""

import os
import re
import base64
import time
from pathlib import Path
from typing import Optional, Dict, Any
from abc import ABC, abstractmethod


# 系统提示词 - 定义 AI 如何生成 SVG
SYSTEM_PROMPT = """你是一个专业的 SVG 生成专家。将提供的架构图转换为干净、结构化的 SVG 代码。

要求：
1. 保留所有文本内容和相对位置
2. 使用矩形(<rect>)、圆形(<circle>)、文本(<text>)、线条(<line>)、多边形(<polygon>)等基础元素
3. 颜色使用 rgb() 或十六进制格式，不要使用 CSS 类
4. 设置合适的 viewBox 尺寸，确保内容完整显示
5. 使用 <g> 分组相关元素
6. 文字使用常见中文字体如 "Microsoft YaHei", "SimHei", "SimSun"
7. 确保 SVG 可以在 PowerPoint 中正常显示
8. 输出纯 SVG 代码，不要 markdown 代码块，不要任何解释文字

布局分析（关键步骤）：
在生成SVG之前，请先分析原图的布局结构：
- 识别主要区域和分组（如左/中/右栏，上/中/下层）
- 分析元素间的对齐关系：哪些元素是水平对齐的（上边缘/下边缘/中心线对齐），哪些是垂直对齐的（左边缘/右边缘/中心线对齐）
- 识别连接关系：哪些元素之间有连线或箭头连接
- 确定间距模式：同类元素之间的标准间距是多少

布局要求（非常重要）：
- 相邻元素之间保持至少 15-20 像素的间距
- 左右并列的元素之间要有明显间隔，不要贴在一起
- 上下排列的元素之间要有足够的留白
- 严格保持原图的对齐关系：水平对齐的元素y坐标要一致，垂直对齐的元素x坐标要一致
- 确保所有元素不重叠、不遮挡
- 连接线的端点要对齐元素中心或边缘

箭头要求：
- 使用 <defs> 定义可复用的 arrowhead marker
- 所有带箭头的线条必须使用 marker-end="url(#arrowhead)" 引用
- 双向箭头使用 marker-start 和 marker-end 同时标记
- 确保箭头方向正确指向目标元素

完整性要求（至关重要）：
- 必须生成完整的 SVG 代码，包括 </svg> 结束标签
- 所有 <g> 分组必须有对应的 </g> 闭合标签
- 所有 XML 元素必须正确闭合
- 不要截断输出，确保整个 SVG 结构完整

SVG 格式要求：
- 包含 xmlns="http://www.w3.org/2000/svg" 命名空间
- 文字元素使用 <text> 标签，设置 font-family、font-size、fill 属性
- 矩形使用 <rect> 标签，设置 x, y, width, height, fill, stroke 等属性
- 直线使用 <line> 标签，可添加箭头标记
- **折线（直角转弯的连接线）必须使用 <polyline> 标签或 <path> 的 L 命令，不要用 Q/C 曲线命令**
- **原图中的折线（如根因定界报告那条线）是先水平后垂直的直角线，不是曲线**
- 确保所有闭合标签正确"""


class BaseAIConverter(ABC):
    """AI 转换器基类"""

    @abstractmethod
    def convert(self, image_path: str, output_path: Optional[str] = None) -> str:
        """将图片转换为 SVG"""
        pass

    def _encode_image(self, image_path: Path) -> tuple:
        """
        将图片编码为 base64，如果太大则自动压缩
        Returns: (base64_data, mime_type)
        """
        from PIL import Image
        import io

        # 读取原始图片
        with Image.open(image_path) as img:
            # 转换为RGB模式（处理PNG透明通道等）
            if img.mode in ('RGBA', 'LA', 'P'):
                img = img.convert('RGB')

            # 检查文件大小，如果太大则压缩
            file_size = image_path.stat().st_size
            max_size = 15 * 1024 * 1024  # 15MB原始文件（base64后会变成约20MB）

            if file_size > max_size:
                # 计算缩放比例
                scale = (max_size / file_size) ** 0.5
                new_width = int(img.width * scale)
                new_height = int(img.height * scale)
                img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
                print(f"  图片过大({file_size/1024/1024:.1f}MB)，已缩放至 {new_width}x{new_height}")

            # 保存为JPEG（通常比PNG小）
            buffer = io.BytesIO()
            img.save(buffer, format='JPEG', quality=85, optimize=True)
            image_data = buffer.getvalue()

            # 如果还是太大，进一步降低质量
            while len(image_data) > max_size and buffer:
                buffer = io.BytesIO()
                img.save(buffer, format='JPEG', quality=60, optimize=True)
                image_data = buffer.getvalue()
                print(f"  进一步压缩图片至 {len(image_data)/1024/1024:.1f}MB")
                break  # 只压缩一次避免无限循环

        mime_type = "image/jpeg"
        return base64.b64encode(image_data).decode('utf-8'), mime_type

    def _clean_svg(self, content: str) -> str:
        """清理 AI 输出的 SVG 内容"""
        # 移除 markdown 代码块标记
        content = re.sub(r'^```svg\s*', '', content, flags=re.IGNORECASE)
        content = re.sub(r'^```\s*', '', content)
        content = re.sub(r'```\s*$', '', content)
        content = content.strip()

        # 确保有 XML 声明
        if not content.startswith('<?xml'):
            content = '<?xml version="1.0" encoding="UTF-8"?>\n' + content

        # 确保有 SVG 命名空间
        if 'xmlns="http://www.w3.org/2000/svg"' not in content:
            content = content.replace('<svg', '<svg xmlns="http://www.w3.org/2000/svg"', 1)

        return content


class ClaudeConverter(BaseAIConverter):
    """Claude (Anthropic) 转换器 - 支持自定义 base_url，兼容OpenAI SDK"""

    def __init__(self, api_key: str, model: str = "claude-3-5-sonnet-20241022", base_url: str = None):
        self.api_key = api_key
        self.model = model
        self.base_url = base_url
        self.use_openai_sdk = False

        try:
            import anthropic
            if base_url:
                self.client = anthropic.Anthropic(api_key=api_key, base_url=base_url)
            else:
                self.client = anthropic.Anthropic(api_key=api_key)
        except ImportError:
            # 如果没有anthropic SDK，尝试使用OpenAI SDK（兼容模式）
            try:
                import openai
                self.client = openai.OpenAI(api_key=api_key, base_url=base_url)
                self.use_openai_sdk = True
                print("  [INFO] 使用OpenAI SDK兼容模式")
            except ImportError:
                raise ImportError("请安装 anthropic SDK: pip install anthropic 或 openai SDK: pip install openai")

    def convert(self, image_path: str, output_path: Optional[str] = None,
                max_retries: int = 3, retry_delay: float = 2.0) -> str:
        image_path = Path(image_path)
        if not image_path.exists():
            raise FileNotFoundError(f"图片不存在: {image_path}")

        if output_path is None:
            output_path = image_path.with_suffix('.svg')
        else:
            output_path = Path(output_path)

        print(f"[PNG->SVG] 转换: {image_path.name}")
        provider_name = "Claude" if not self.base_url else f"Claude-compatible ({self.base_url[:30]}...)"
        sdk_name = "OpenAI SDK" if self.use_openai_sdk else "Anthropic SDK"
        print(f"  使用模型: {provider_name} / {self.model} ({sdk_name})")

        image_data, mime_type = self._encode_image(image_path)

        for attempt in range(max_retries):
            try:
                if self.use_openai_sdk:
                    # 使用OpenAI SDK兼容模式
                    image_url = f"data:{mime_type};base64,{image_data}"
                    response = self.client.chat.completions.create(
                        model=self.model,
                        messages=[
                            {
                                "role": "system",
                                "content": SYSTEM_PROMPT
                            },
                            {
                                "role": "user",
                                "content": [
                                    {
                                        "type": "image_url",
                                        "image_url": {"url": image_url}
                                    },
                                    {
                                        "type": "text",
                                        "text": "将这张架构图转换为 SVG 代码。保留所有文本、颜色、布局和元素样式。直接输出 SVG 代码，不要添加 markdown 标记或其他说明文字。"
                                    }
                                ]
                            }
                        ],
                        temperature=0.2,
                        max_tokens=8192
                    )
                    svg_content = self._clean_svg(response.choices[0].message.content)
                else:
                    # 使用原生Anthropic SDK
                    response = self.client.messages.create(
                        model=self.model,
                        max_tokens=8192,
                        temperature=0.2,
                        system=SYSTEM_PROMPT,
                        messages=[{
                            "role": "user",
                            "content": [
                                {
                                    "type": "image",
                                    "source": {
                                        "type": "base64",
                                        "media_type": mime_type,
                                        "data": image_data
                                    }
                                },
                                {
                                    "type": "text",
                                    "text": "将这张架构图转换为 SVG 代码。保留所有文本、颜色、布局和元素样式。直接输出 SVG 代码，不要添加 markdown 标记或其他说明文字。"
                                }
                            ]
                        }]
                    )
                    # 处理不同类型的 content block
                    svg_content = None
                    for block in response.content:
                        if hasattr(block, 'text'):
                            svg_content = self._clean_svg(block.text)
                            break
                    if svg_content is None:
                        raise RuntimeError(f"响应中没有找到文本内容，content types: {[type(b).__name__ for b in response.content]}")

                output_path.write_text(svg_content, encoding='utf-8')
                print(f"  [OK] 成功生成: {output_path}")
                return str(output_path)

            except Exception as e:
                print(f"  [FAIL] 尝试 {attempt + 1}/{max_retries} 失败: {e}")
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                else:
                    raise RuntimeError(f"转换失败: {e}")


class KimiConverter(BaseAIConverter):
    """Kimi (Moonshot AI) 转换器 - 国内可用"""

    def __init__(self, api_key: str, model: str = "moonshot-v1-32k-vision-preview"):
        self.api_key = api_key
        self.model = model
        self.base_url = "https://api.moonshot.cn/v1"

        try:
            import openai
            self.client = openai.OpenAI(
                api_key=api_key,
                base_url=self.base_url
            )
        except ImportError:
            raise ImportError("请安装 openai SDK: pip install openai")

    def convert(self, image_path: str, output_path: Optional[str] = None,
                max_retries: int = 3, retry_delay: float = 2.0) -> str:
        image_path = Path(image_path)
        if not image_path.exists():
            raise FileNotFoundError(f"图片不存在: {image_path}")

        if output_path is None:
            output_path = image_path.with_suffix('.svg')
        else:
            output_path = Path(output_path)

        print(f"[PNG->SVG] 转换: {image_path.name}")
        print(f"  使用模型: Kimi ({self.model})")

        image_data, mime_type = self._encode_image(image_path)
        image_url = f"data:{mime_type};base64,{image_data}"

        for attempt in range(max_retries):
            try:
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=[
                        {
                            "role": "system",
                            "content": SYSTEM_PROMPT
                        },
                        {
                            "role": "user",
                            "content": [
                                {
                                    "type": "image_url",
                                    "image_url": {"url": image_url}
                                },
                                {
                                    "type": "text",
                                    "text": "将这张架构图转换为 SVG 代码。保留所有文本、颜色、布局和元素样式。直接输出 SVG 代码，不要添加 markdown 标记或其他说明文字。"
                                }
                            ]
                        }
                    ],
                    temperature=0.2,
                    max_tokens=4096
                )

                svg_content = self._clean_svg(response.choices[0].message.content)
                output_path.write_text(svg_content, encoding='utf-8')
                print(f"  [OK] 成功生成: {output_path}")
                return str(output_path)

            except Exception as e:
                print(f"  [FAIL] 尝试 {attempt + 1}/{max_retries} 失败: {e}")
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                else:
                    raise RuntimeError(f"转换失败: {e}")


class GLMConverter(BaseAIConverter):
    """智谱 GLM 转换器 - 国内可用"""

    def __init__(self, api_key: str, model: str = "glm-4v"):
        self.api_key = api_key
        self.model = model
        self.base_url = "https://open.bigmodel.cn/api/paas/v4"

        try:
            import openai
            self.client = openai.OpenAI(
                api_key=api_key,
                base_url=self.base_url
            )
        except ImportError:
            raise ImportError("请安装 openai SDK: pip install openai")

    def convert(self, image_path: str, output_path: Optional[str] = None,
                max_retries: int = 3, retry_delay: float = 2.0) -> str:
        image_path = Path(image_path)
        if not image_path.exists():
            raise FileNotFoundError(f"图片不存在: {image_path}")

        if output_path is None:
            output_path = image_path.with_suffix('.svg')
        else:
            output_path = Path(output_path)

        print(f"[PNG->SVG] 转换: {image_path.name}")
        print(f"  使用模型: GLM ({self.model})")

        image_data, mime_type = self._encode_image(image_path)
        image_url = f"data:{mime_type};base64,{image_data}"

        for attempt in range(max_retries):
            try:
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=[
                        {
                            "role": "system",
                            "content": SYSTEM_PROMPT
                        },
                        {
                            "role": "user",
                            "content": [
                                {
                                    "type": "image_url",
                                    "image_url": {"url": image_url}
                                },
                                {
                                    "type": "text",
                                    "text": "将这张架构图转换为 SVG 代码。保留所有文本、颜色、布局和元素样式。直接输出 SVG 代码，不要添加 markdown 标记或其他说明文字。"
                                }
                            ]
                        }
                    ],
                    temperature=0.2,
                    max_tokens=4096
                )

                svg_content = self._clean_svg(response.choices[0].message.content)
                output_path.write_text(svg_content, encoding='utf-8')
                print(f"  [OK] 成功生成: {output_path}")
                return str(output_path)

            except Exception as e:
                print(f"  [FAIL] 尝试 {attempt + 1}/{max_retries} 失败: {e}")
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                else:
                    raise RuntimeError(f"转换失败: {e}")


class OpenAIConverter(BaseAIConverter):
    """OpenAI GPT-4V 转换器"""

    def __init__(self, api_key: str, model: str = "gpt-4o", base_url: str = None):
        self.api_key = api_key
        self.model = model
        self.base_url = base_url

        try:
            import openai
            if base_url:
                self.client = openai.OpenAI(api_key=api_key, base_url=base_url)
            else:
                self.client = openai.OpenAI(api_key=api_key)
        except ImportError:
            raise ImportError("请安装 openai SDK: pip install openai")

    def convert(self, image_path: str, output_path: Optional[str] = None,
                max_retries: int = 3, retry_delay: float = 2.0) -> str:
        image_path = Path(image_path)
        if not image_path.exists():
            raise FileNotFoundError(f"图片不存在: {image_path}")

        if output_path is None:
            output_path = image_path.with_suffix('.svg')
        else:
            output_path = Path(output_path)

        print(f"[PNG->SVG] 转换: {image_path.name}")
        print(f"  使用模型: OpenAI ({self.model})")

        image_data, mime_type = self._encode_image(image_path)
        image_url = f"data:{mime_type};base64,{image_data}"

        for attempt in range(max_retries):
            try:
                response = self.client.chat.completions.create(
                    model=self.model,
                    messages=[
                        {
                            "role": "system",
                            "content": SYSTEM_PROMPT
                        },
                        {
                            "role": "user",
                            "content": [
                                {
                                    "type": "image_url",
                                    "image_url": {"url": image_url}
                                },
                                {
                                    "type": "text",
                                    "text": "将这张架构图转换为 SVG 代码。保留所有文本、颜色、布局和元素样式。直接输出 SVG 代码，不要添加 markdown 标记或其他说明文字。"
                                }
                            ]
                        }
                    ],
                    temperature=0.2,
                    max_tokens=4096
                )

                svg_content = self._clean_svg(response.choices[0].message.content)
                output_path.write_text(svg_content, encoding='utf-8')
                print(f"  [OK] 成功生成: {output_path}")
                return str(output_path)

            except Exception as e:
                print(f"  [FAIL] 尝试 {attempt + 1}/{max_retries} 失败: {e}")
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                else:
                    raise RuntimeError(f"转换失败: {e}")


class MockAIConverter(BaseAIConverter):
    """模拟 AI 转换器，用于测试（不调用 API）"""

    def convert(self, image_path: str, output_path: Optional[str] = None) -> str:
        """生成一个简单的测试 SVG"""
        image_path = Path(image_path)
        if output_path is None:
            output_path = image_path.with_suffix('.svg')
        else:
            output_path = Path(output_path)

        svg_content = '''<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="800" height="600" viewBox="0 0 800 600">
  <rect x="10" y="10" width="780" height="580" fill="none" stroke="red" stroke-width="2"/>
  <text x="400" y="300" font-family="Microsoft YaHei" font-size="24" text-anchor="middle" fill="red">
    [测试] 这是一个模拟生成的 SVG
请在设置 API Key 后使用真实转换
  </text>
</svg>'''

        output_path.write_text(svg_content, encoding='utf-8')
        print(f"[MOCK] 生成测试 SVG: {output_path}")
        return str(output_path)


class AIImageToSVGConverter:
    """
    统一的 AI 图片转 SVG 转换器
    自动根据环境变量选择后端
    """

    PROVIDERS = {
        'claude': ClaudeConverter,
        'kimi': KimiConverter,
        'glm': GLMConverter,
        'openai': OpenAIConverter,
    }

    def __init__(self, provider: str = None, api_key: str = None, model: str = None, base_url: str = None):
        """
        初始化转换器

        Args:
            provider: AI 提供商 ('claude', 'kimi', 'glm', 'openai')
                     默认从环境变量 AI_PROVIDER 读取，否则自动检测
            api_key: API Key，默认从对应环境变量读取
            model: 模型名称，默认使用提供商推荐模型
            base_url: 自定义 API base URL（用于兼容API代理）
        """
        # 确定提供商
        self.provider = provider or os.environ.get('AI_PROVIDER', 'auto')

        if self.provider == 'auto':
            self.provider = self._auto_detect_provider()

        if self.provider not in self.PROVIDERS:
            raise ValueError(f"不支持的提供商: {self.provider}。支持: {list(self.PROVIDERS.keys())}")

        # 获取 API Key
        self.api_key = api_key or self._get_api_key(self.provider)

        if not self.api_key:
            raise ValueError(
                f"请提供 API Key 或设置对应环境变量:\n"
                f"  Claude: ANTHROPIC_API_KEY\n"
                f"  Kimi: MOONSHOT_API_KEY\n"
                f"  GLM: ZHIPU_API_KEY\n"
                f"  OpenAI: OPENAI_API_KEY"
            )

        # 创建对应转换器
        converter_class = self.PROVIDERS[self.provider]

        # ClaudeConverter 支持 base_url 参数
        if self.provider == 'claude' and base_url:
            self.converter = converter_class(self.api_key, model or "claude-3-5-sonnet-20241022", base_url)
        elif model:
            self.converter = converter_class(self.api_key, model)
        else:
            self.converter = converter_class(self.api_key)

    def _auto_detect_provider(self) -> str:
        """根据环境变量自动检测提供商"""
        if os.environ.get('ANTHROPIC_API_KEY'):
            return 'claude'
        elif os.environ.get('MOONSHOT_API_KEY'):
            return 'kimi'
        elif os.environ.get('ZHIPU_API_KEY'):
            return 'glm'
        elif os.environ.get('OPENAI_API_KEY'):
            return 'openai'
        else:
            return 'claude'  # 默认

    def _get_api_key(self, provider: str) -> Optional[str]:
        """获取对应提供商的 API Key"""
        env_vars = {
            'claude': 'ANTHROPIC_API_KEY',
            'kimi': 'MOONSHOT_API_KEY',
            'glm': 'ZHIPU_API_KEY',
            'openai': 'OPENAI_API_KEY',
        }
        return os.environ.get(env_vars.get(provider))

    def convert(self, image_path: str, output_path: Optional[str] = None, **kwargs) -> str:
        """转换图片为 SVG"""
        return self.converter.convert(image_path, output_path, **kwargs)


# 保持向后兼容
__all__ = [
    'AIImageToSVGConverter',
    'ClaudeConverter',
    'KimiConverter',
    'GLMConverter',
    'OpenAIConverter',
    'MockAIConverter',
]


if __name__ == "__main__":
    # 测试代码
    import sys

    if len(sys.argv) < 2:
        print("用法: python ai_client.py <图片路径> [--provider claude|kimi|glm|openai]")
        sys.exit(1)

    image_file = sys.argv[1]

    # 解析参数
    provider = None
    if '--provider' in sys.argv:
        idx = sys.argv.index('--provider')
        provider = sys.argv[idx + 1] if idx + 1 < len(sys.argv) else None

    # 检查是否有 API Key
    try:
        converter = AIImageToSVGConverter(provider=provider)
    except ValueError as e:
        print(f"警告: {e}")
        print("使用模拟模式")
        converter = MockAIConverter()

    try:
        result = converter.convert(image_file)
        print(f"\n转换完成: {result}")
    except Exception as e:
        print(f"\n错误: {e}")
        sys.exit(1)
