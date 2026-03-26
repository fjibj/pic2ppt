# Pic2PPT - AI图片转PPT可编辑形状

将AI生成的架构图（SVG/PNG格式）转换为PowerPoint中可编辑的形状。

## 功能特性

- ✅ SVG → PPTX 转换
- ✅ PNG图片智能识别转PPT
- ✅ 保留矩形、文本、线条等基础形状
- ✅ 保留颜色和样式
- ✅ 可编辑的文本（不是图片）
- ✅ 绿色便携版，单文件EXE，双击即运行
- ✅ 支持文件选择、拖拽、粘贴三种上传方式

## 安装依赖

```bash
pip install -r requirements.txt
```

依赖项：
- python-pptx >= 0.6.21
- Pillow >= 9.0.0
- lxml >= 4.9.0
- anthropic (用于AI识别)
- openai (用于AI识别)

## 使用方法

### 方式1：直接运行 Python 脚本

```bash
python pic2ppt.py
```

### 方式2：运行 SVG 转换脚本

```bash
python src/svg_to_pptx.py
```

默认会读取 `ai_component_management_platform.svg` 并生成 `output_architecture.pptx`。

### 方式3：在 Python 代码中使用

```python
from src.svg_to_pptx import SVGToPPTXConverter

converter = SVGToPPTXConverter("your_diagram.svg")
converter.convert("output.pptx", slide_width=13.333, slide_height=10)
```

## 支持转换的SVG元素

| SVG元素 | PPT形状 | 说明 |
|---------|---------|------|
| `<rect>` | 矩形/圆角矩形 | 支持填充色和边框 |
| `<text>` | 文本框 | 保留字体、颜色和对齐方式 |
| `<line>` | 直线/箭头 | 支持颜色和线宽 |

## 注意事项

- 仅支持Microsoft PowerPoint（WPS不支持"转换为形状"功能）
- 复杂的SVG渐变效果可能无法完全保留
- 建议在PowerPoint中打开后微调布局
- **中文支持**：文本使用微软雅黑字体

## 文件说明

- `pic2ppt.py` - 主GUI程序（支持PNG图片识别）
- `src/svg_to_pptx.py` - SVG转换核心程序
- `src/usage_limiter.py` - 使用限制管理
- `requirements.txt` - Python依赖列表
- `CLAUDE.md` - 开发文档
- `QUICKSTART.md` - 快速开始指南

## 项目结构

```
pic2shape/
├── pic2ppt.py              # 主程序入口
├── src/                    # 源代码目录
│   ├── svg_to_pptx.py     # SVG转PPTX核心
│   ├── usage_limiter.py   # 使用限制
│   └── ...
├── docs/                   # 文档目录
├── dist/                   # 打包输出
├── requirements.txt        # 依赖列表
├── README.md              # 项目说明
└── CLAUDE.md              # 开发文档
```

## License

MIT
