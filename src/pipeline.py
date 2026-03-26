#!/usr/bin/env python3
"""
PNG to PPTX Pipeline
完整的图片到 PowerPoint 转换流程
支持多种 AI 提供商
"""

import os
import sys
import logging
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

# 添加项目路径
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from png2svg import AIImageToSVGConverter, SVGValidator
from svg_converter import SVGToPPTXConverter


class PNGToPPTXPipeline:
    """
    完整流程: PNG → SVG → PPTX
    """

    def __init__(self, provider: str = None, api_key: str = None, base_url: str = None, model: str = None):
        """
        初始化流水线

        Args:
            provider: AI 提供商 ('claude', 'kimi', 'glm', 'openai')
            api_key: API Key，默认从对应环境变量读取
            base_url: API Base URL，可选
            model: AI 模型名称，可选
        """
        self.provider = provider
        self.api_key = api_key
        self.base_url = base_url
        self.model = model
        self.progress_callback = None

    def set_progress_callback(self, callback):
        """
        设置进度回调函数

        Args:
            callback: 回调函数，参数为 (step, total, message)
        """
        self.progress_callback = callback

    def _notify_progress(self, step: int, total: int, message: str):
        """通知进度更新"""
        if self.progress_callback:
            try:
                self.progress_callback(step, total, message)
            except Exception as e:
                logger.error(f"进度回调出错: {e}")

    def convert(self, image_path: str, output_pptx: Optional[str] = None,
                keep_svg: bool = True) -> str:
        """
        将图片转换为 PowerPoint

        Args:
            image_path: 输入图片路径 (PNG/JPG)
            output_pptx: 输出 PPTX 路径，默认与输入同名
            keep_svg: 是否保留中间 SVG 文件

        Returns:
            生成的 PPTX 文件路径
        """
        image_path = Path(image_path)

        if not image_path.exists():
            raise FileNotFoundError(f"图片不存在: {image_path}")

        # 检查文件类型
        if image_path.suffix.lower() not in ['.png', '.jpg', '.jpeg', '.bmp', '.gif', '.webp', '.tiff', '.tif']:
            raise ValueError(f"不支持的文件格式: {image_path.suffix}")

        # 设置输出路径
        if output_pptx is None:
            output_pptx = image_path.with_suffix('.pptx')
        else:
            output_pptx = Path(output_pptx)

        svg_path = image_path.with_suffix('.svg')

        print("=" * 70)
        print("PNG to PPTX 转换流程")
        print("=" * 70)
        print(f"输入: {image_path}")
        print(f"输出: {output_pptx}")
        print()

        try:
            # 步骤 1: PNG → SVG
            print("[步骤 1/2] PNG -> SVG 转换")
            print("-" * 40)
            self._notify_progress(1, 3, "正在初始化 AI 客户端...")

            png_converter = AIImageToSVGConverter(
                provider=self.provider,
                api_key=self.api_key,
                base_url=self.base_url,
                model=self.model
            )
            self._notify_progress(2, 3, "AI 正在生成 SVG，请稍候...")
            svg_file = png_converter.convert(str(image_path), str(svg_path))

            # 验证 SVG
            self._notify_progress(3, 3, "正在验证 SVG...")
            validator = SVGValidator()
            svg_content = Path(svg_file).read_text(encoding='utf-8')
            is_valid, errors = validator.validate(svg_content)

            if not is_valid:
                print(f"  警告: SVG 验证发现问题:")
                for err in errors:
                    print(f"    - {err}")
                print("  尝试自动修复...")
                svg_content = validator.fix_common_issues(svg_content)
                Path(svg_file).write_text(svg_content, encoding='utf-8')
                print("  [OK] 已修复")

            stats = validator.get_statistics(svg_content)
            print(f"  SVG 大小: {stats['total_length']} 字节")
            print(f"  元素数量: {sum(stats['element_counts'].values())}")
            print()

            # 步骤 2: SVG → PPTX
            print("[步骤 2/2] SVG -> PPTX 转换")
            print("-" * 40)

            pptx_converter = SVGToPPTXConverter(svg_file)
            result = pptx_converter.convert(str(output_pptx))

            print()
            print("=" * 70)
            print("[OK] 转换成功!")
            print(f"  PPTX: {result}")
            if keep_svg:
                print(f"  SVG:  {svg_file}")
            else:
                os.remove(svg_file)
                print(f"  SVG:  已删除")
            print("=" * 70)

            return result

        except Exception as e:
            print()
            print("=" * 70)
            print(f"[FAIL] 转换失败: {e}")
            print("=" * 70)
            raise


def main():
    """命令行入口"""
    import argparse

    parser = argparse.ArgumentParser(
        description='将 PNG/JPG 图片转换为 PowerPoint 可编辑形状',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python pipeline.py diagram.png
  python pipeline.py diagram.png -o output.pptx
  python pipeline.py diagram.png --no-keep-svg

支持的 AI 提供商:
  claude  - Claude (Anthropic) 默认
  kimi    - Kimi (Moonshot AI) 国内可用
  glm     - 智谱 GLM 国内可用
  openai  - OpenAI GPT-4V

环境变量:
  ANTHROPIC_API_KEY  - Claude API Key
  MOONSHOT_API_KEY   - Kimi API Key
  ZHIPU_API_KEY      - GLM API Key
  OPENAI_API_KEY     - OpenAI API Key
        """
    )

    parser.add_argument('image', help='输入图片路径 (PNG/JPG)')
    parser.add_argument('-o', '--output', help='输出 PPTX 路径')
    parser.add_argument('--provider', choices=['claude', 'kimi', 'glm', 'openai'],
                        help='AI 提供商 (默认自动检测)')
    parser.add_argument('--no-keep-svg', action='store_true',
                        help='不保留中间 SVG 文件')
    parser.add_argument('--api-key', help='API Key (默认从环境变量读取)')

    args = parser.parse_args()

    try:
        pipeline = PNGToPPTXPipeline(provider=args.provider, api_key=args.api_key)
        result = pipeline.convert(
            args.image,
            output_pptx=args.output,
            keep_svg=not args.no_keep_svg
        )
        return 0
    except Exception as e:
        print(f"\n错误: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
