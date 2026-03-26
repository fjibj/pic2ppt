# Pic2Shape - AI图片转PPT可编辑形状

## 项目概述

本项目支持完整的图片到PPT转换流程：
1. **PNG/JPG → SVG**：使用AI视觉模型自动识别图片中的元素、文字、布局
2. **SVG → PPTX**：将SVG转换为PowerPoint中可编辑的形状

支持完整的元素识别：矩形、圆形、椭圆、线条、多边形、路径、文本、分组等。

## Claude Code Skills

### /svg2ppt

将SVG/PNG/JPG文件转换为PowerPoint可编辑形状。

**用法：**
```
/svg2ppt <文件路径>              # 转换单个文件
/svg2ppt <通配符路径>            # 批量转换（支持 *.svg, *.png）
```

**参数：**
- `文件路径` - 支持的文件格式：.svg, .png, .jpg, .jpeg
- `通配符路径` - 支持 `*` 和 `?` 通配符

**支持的AI提供商：**
| 提供商 | 环境变量 | 说明 |
|--------|----------|------|
| Claude | `ANTHROPIC_API_KEY` | Anthropic官方，默认 |
| Kimi | `MOONSHOT_API_KEY` | Moonshot AI，国内可用 |
| GLM | `ZHIPU_API_KEY` | 智谱AI，国内可用 |
| OpenAI | `OPENAI_API_KEY` | GPT-4V |

**示例：**
```
/svg2ppt diagram.svg                              # SVG转PPTX
/svg2ppt diagram.png                              # PNG转PPTX（自动检测AI）
/svg2ppt *.svg                                    # 批量转换SVG
/svg2ppt *.png                                    # 批量转换PNG
/svg2ppt F:\projects\diagrams\*.svg            # 指定目录批量转换
```

**输出：**
- 生成的PPT文件保存在同目录
- 文件名与原文件相同，扩展名改为 `.pptx`
- 例如：`diagram.png` → `diagram.pptx`（中间生成`diagram.svg`）

**环境变量（自动检测优先级）：**
```
ANTHROPIC_API_KEY  - Claude API Key
MOONSHOT_API_KEY   - Kimi API Key（国内）
ZHIPU_API_KEY      - 智谱 GLM API Key（国内）
OPENAI_API_KEY     - OpenAI API Key
```

设置环境变量后，程序会自动检测可用的AI提供商。

**支持的SVG元素：**
| SVG元素 | PPT形状 |
|---------|---------|
| `<rect>` | 矩形/圆角矩形 |
| `<circle>` / `<ellipse>` | 椭圆 |
| `<line>` / `<path>`(直线) | 直线连接器（支持箭头） |
| `<polygon>` | 自动识别：菱形/六边形/三角形等 |
| `<text>` | 文本框 |
| `<g>` | 分组（支持样式继承） |

## 完整流程

### 流程1：已有SVG → PPTX（推荐）
```
diagram.svg → [svg_converter] → diagram.pptx
```

### 流程2：PNG/JPG → SVG → PPTX（AI生成）
```
diagram.png → [AI Vision] → diagram.svg → [svg_converter] → diagram.pptx
```

**AI生成说明：**
- 使用Claude/GPT-4V等视觉模型分析图片
- 自动识别：元素位置、形状、颜色、文字内容
- 生成高质量SVG，保留原图布局和样式

## 技术栈

- Python 3.7+
- PPT生成：python-pptx
- SVG解析：lxml
- AI视觉：Claude/Kimi/GLM/OpenAI API (PNG/JPG转SVG)

## 支持的AI提供商

| 提供商 | 模型 | 国内可用 | 成本 | 质量 |
|--------|------|----------|------|------|
| Claude | claude-3-5-sonnet | 否 | 中 | 极高 |
| Kimi | moonshot-v1-32k-vision | 是 | 低 | 高 |
| GLM | glm-4v | 是 | 低 | 高 |
| OpenAI | gpt-4o | 否 | 高 | 极高 |

## 使用方法

### GUI应用（推荐）

**启动方式**:
```bash
# 方式1：直接运行（开发环境）
python pic2ppt.py

# 方式2：双击EXE（打包版本）
dist/pic2ppt.exe
```

**功能说明**:
1. **图片上传**：支持拖拽、Ctrl+V粘贴、点击选择
2. **API配置**：支持Kimi、Claude、OpenAI、GLM
3. **输出设置**：自定义输出路径
4. **转换**：一键转换图片为可编辑PPT

**界面布局**:
```
┌─────────────────────────────────────┐
│  pic2ppt v1.0 - AI图片转PPT        │
├─────────────────────────────────────┤
│  ┌─────────────────────────────┐   │
│  │      图片预览区              │   │  ← 支持拖拽/粘贴
│  │   (拖拽或Ctrl+V粘贴图片)     │   │
│  └─────────────────────────────┘   │
├─────────────────────────────────────┤
│  ▼ API配置 (可折叠)                 │
│  ├─ Base URL: [____________]        │
│  ├─ API Key:  [____________] [👁]   │
│  ├─ 模型:     [kimi-k2.5___] [测试] │
├─────────────────────────────────────┤
│  ▼ 输出设置 (可折叠)                 │
│  └─ 输出路径: [____________] [浏览] │
├─────────────────────────────────────┤
│  [转换] [取消]        [进度条]       │
├─────────────────────────────────────┤
│  状态: 就绪                          │
└─────────────────────────────────────┘
```

**支持的AI配置示例**:

| 提供商 | Base URL | 模型 |
|--------|----------|------|
| Kimi | `https://api.moonshot.cn/v1` | `moonshot-v1-32k-vision` |
| Kimi(DashScope) | `https://coding.dashscope.aliyuncs.com/apps/anthropic` | `kimi-k2.5` |
| Claude | `https://api.anthropic.com` | `claude-3-5-sonnet-20241022` |
| OpenAI | `https://api.openai.com/v1` | `gpt-4o` |
| GLM | `https://open.bigmodel.cn/api/paas/v4` | `glm-4v` |

### 命令行工具

**SVG → PPTX：**
```bash
# 转换单个文件
python src/convert.py diagram.svg

# 指定输出文件
python src/convert.py diagram.svg -o output.pptx

# 指定幻灯片尺寸
python src/convert.py diagram.svg -W 16 -H 9
```

**PNG/JPG → PPTX：**
```bash
# 设置环境变量（选择其中一个）
set ANTHROPIC_API_KEY=your_key_here    # Claude
set MOONSHOT_API_KEY=your_key_here     # Kimi（国内）
set ZHIPU_API_KEY=your_key_here        # 智谱 GLM（国内）

# 转换单个文件（自动检测AI提供商）
python src/convert.py diagram.png

# 指定AI提供商
python src/convert.py diagram.png --provider kimi
python src/convert.py diagram.png --provider glm

# 手动指定API Key
python src/convert.py diagram.png --api-key sk-xxx --provider kimi

# 不保留中间SVG文件
python src/convert.py diagram.png --no-keep-svg
```

**批量转换：**
```bash
# 批量转换SVG
python src/pic2shape.py *.svg

# 批量转换PNG
python src/pic2shape.py *.png

# 混合批量转换
python src/pic2shape.py *.svg *.png
```

### Claude Code Skill

使用 `/svg2ppt` 命令快速转换：

```
/svg2ppt diagram.svg                    # 转换SVG
/svg2ppt diagram.png                    # 转换PNG（自动AI识别）
/svg2ppt *.svg                          # 批量转换SVG
/svg2ppt *.png                          # 批量转换PNG
/svg2ppt "F:\path\to\*.svg"            # 指定目录
```

## 支持的SVG元素

| 元素 | 映射到PPT形状 |
|------|--------------|
| `<rect>` | 矩形/圆角矩形 |
| `<circle>` | 椭圆 |
| `<ellipse>` | 椭圆 |
| `<line>` | 直线连接器 |
| `<polygon>` | 自动识别为菱形/六边形等 |
| `<path>` | 直线带箭头/矩形 |
| `<text>` | 文本框 |
| `<g>` | 分组（样式继承） |

## 文件结构

```
pic2shape/
├── CLAUDE.md                          # 项目文档
├── 应用打包方案.md                     # 打包配置文档
├── pic2ppt.py                         # GUI主程序
├── pic2ppt.spec                       # PyInstaller多文件配置
├── pic2ppt_onefile.spec               # PyInstaller单文件配置
├── config.json                        # 用户配置（自动生成）
├── pic2ppt.log                        # 运行日志（自动生成）
│
├── src/                               # 核心模块
│   ├── __init__.py
│   ├── convert.py                     # 命令行工具
│   ├── pic2shape.py                   # 批量转换工具
│   ├── pipeline.py                    # 完整转换流程
│   ├── png2svg/                       # PNG→SVG模块
│   │   ├── __init__.py
│   │   ├── ai_client.py               # AI API客户端
│   │   └── validator.py               # SVG验证器
│   └── svg_converter/                 # SVG→PPT模块
│       ├── __init__.py
│       ├── converter.py               # 主转换器
│       ├── parser.py                  # SVG解析器
│       ├── models.py                  # 数据模型
│       ├── geometry.py                # 几何分析引擎
│       ├── color_utils.py             # 颜色解析
│       └── handlers/                  # 元素处理器
│           ├── __init__.py
│           ├── basic_shapes.py
│           ├── polygons.py
│           ├── text.py
│           └── group.py
│
├── dist/                              # 打包输出
│   ├── pic2ppt.exe                    # 单文件EXE (33MB)
│   └── 启动程序.bat                    # 带日志查看的启动脚本
│
└── .claude/commands/
    └── svg2ppt.md                     # Claude Code Skill定义
```

## 注意事项

- WPS PowerPoint不支持"转换为形状"功能，仅Microsoft PowerPoint支持
- PNG/JPG转换需要有效的 API Key（支持Claude、Kimi、GLM、OpenAI）
- **国内用户推荐使用 Kimi 或 智谱 GLM**，无需翻墙
- AI生成的SVG质量取决于原图清晰度
- 建议在转换前确保图片中的文字清晰可读
- 环境变量设置后需要重启终端才能生效
- **图片大小限制**: 超过15MB的图片会自动压缩（保持质量85%→60%）

## 最新更新

### 2026-03-19 GUI界面与图标优化

| 改进项 | 说明 |
|--------|------|
| 应用图标 | 新增智能魔方风格ICO图标（2×2方块，橙→琥珀渐变） |
| SVG图标源文件 | 创建`pic2ppt_icon.svg`，包含图片/PPT符号、转换箭头和"FangJin"文字 |
| 标题栏图标 | 在界面左上角添加60×60像素大图标，使用Canvas绘制SVG风格图案 |
| 窗口高度 | 从820px优化至780px，更紧凑的界面布局 |
| 图片预览区 | 高度从280px调整至260px，平衡预览效果与界面空间 |
| 按钮显示 | 优化布局确保"开始转换"和"取消"按钮完整可见 |

**图标设计说明**：
- 左侧橙色方块（图片符号：山形+太阳）
- 右侧琥珀方块（PPT符号：三条横线）
- 中央白色圆形+橙色箭头（转换意象）
- 底部"FangJin"文字标识

### 2026-03-18 SVG转PPT转换优化

| 问题 | 修复说明 |
|------|----------|
| 白色矩形不显示 | 无明确填充时默认使用白色填充 |
| 粗箭头显示为方块 | 根据线宽动态选择箭头大小（细线=sm, 中线=med, 粗线=lg） |
| 竖排文字倾斜 | 改用`eaVert`东亚竖排属性，而非旋转 |
| 梯形/平行四边形缺失 | 新增到SHAPE_MAP，支持TRAPEZOID和PARALLELOGRAM |
| 路径坐标解析错误 | 改进正则表达式，支持更多坐标格式 |
| 箭头方向错误 | 修复：PPT中`headEnd`=起点箭头, `tailEnd`=终点箭头（与直觉相反） |
| 大对角线干扰 | 路径点提取只取第一个子路径（M命令之间的内容） |
| 图片过大错误 | 超过15MB自动压缩，先尝试85%质量JPEG，如仍超过则60%质量 |

## 打包说明

### 单文件EXE打包

```bash
# 1. 安装依赖
pip install pyinstaller openai anthropic httpx python-pptx lxml pillow tkinterdnd2

# 2. 打包
pyinstaller pic2ppt_onefile.spec --clean --noconfirm

# 3. 输出
dist/pic2ppt.exe  (33MB)
```

**启动脚本** (`启动程序.bat`):
```batch
@echo off
chcp 65001 >nul
echo 正在启动 pic2ppt...
pic2ppt.exe
if errorlevel 1 (
    echo 程序异常退出，请查看 pic2ppt.log
    pause
)
```
