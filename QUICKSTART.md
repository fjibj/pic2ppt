# Pic2Shape 快速开始指南

## 功能概述

Pic2Shape 支持完整的图片到PPT转换流程：

```
PNG/JPG → [AI Vision] → SVG → [Converter] → PPTX
```

## 安装依赖

```bash
# 基础依赖（SVG → PPTX）
pip install python-pptx lxml

# AI 转换依赖（PNG/JPG → SVG）
pip install anthropic openai
```

## 选择AI提供商

支持多个AI视觉模型，国内用户推荐Kimi或智谱GLM：

| 提供商 | 环境变量 | 国内可用 | 特点 |
|--------|----------|----------|------|
| **Kimi** | `MOONSHOT_API_KEY` | ✅ 是 | 性价比高，速度快 |
| **智谱GLM** | `ZHIPU_API_KEY` | ✅ 是 | 中文优化好 |
| Claude | `ANTHROPIC_API_KEY` | ❌ 否 | 质量极高 |
| OpenAI | `OPENAI_API_KEY` | ❌ 否 | GPT-4V |

## 环境配置

### Windows CMD

```bash
# 选择其中一个设置
set MOONSHOT_API_KEY=your_kimi_key_here       # 推荐国内用户
set ZHIPU_API_KEY=your_glm_key_here           # 推荐国内用户
set ANTHROPIC_API_KEY=your_claude_key_here
set OPENAI_API_KEY=your_openai_key_here
```

### Windows PowerShell

```powershell
# 选择其中一个设置
$env:MOONSHOT_API_KEY="your_kimi_key_here"
$env:ZHIPU_API_KEY="your_glm_key_here"
```

### Linux/Mac

```bash
export MOONSHOT_API_KEY=your_kimi_key_here
```

**注意：** 设置后需要重启终端才能生效。

## 使用方法

### 1. SVG → PPTX（已有SVG文件）

```bash
# 单个文件
python src/convert.py diagram.svg

# 批量转换
python src/pic2shape.py *.svg
```

### 2. PNG/JPG → PPTX（AI生成SVG）

#### 自动检测AI提供商

```bash
# 单个文件
python src/convert.py diagram.png

# 批量转换
python src/pic2shape.py *.png
```

#### 指定AI提供商

```bash
# 使用 Kimi（推荐国内用户）
python src/convert.py diagram.png --provider kimi

# 使用 智谱 GLM
python src/convert.py diagram.png --provider glm

# 使用 Claude
python src/convert.py diagram.png --provider claude
```

#### 手动指定API Key

```bash
# 不设置环境变量，直接在命令行指定
python src/convert.py diagram.png --api-key your_key_here --provider kimi
```

#### 不保留中间SVG

```bash
python src/convert.py diagram.png --no-keep-svg
```

### 3. 混合转换

```bash
# 同时转换SVG和PNG
python src/pic2shape.py *.svg *.png
```

### 4. Claude Code 命令

```bash
# 转换SVG
/svg2ppt diagram.svg

# 转换PNG（自动AI识别）
/svg2ppt diagram.png

# 批量转换
/svg2ppt *.svg
/svg2ppt *.png
```

## 输出文件

- **SVG → PPTX**: `diagram.svg` → `diagram.pptx`
- **PNG → PPTX**: `diagram.png` → `diagram.svg` → `diagram.pptx`

## 获取API Key

### Kimi (Moonshot AI)

1. 访问 https://platform.moonshot.cn/
2. 注册账号
3. 创建 API Key
4. 复制 Key 到环境变量

### 智谱 GLM

1. 访问 https://open.bigmodel.cn/
2. 注册账号
3. 创建应用获取 API Key
4. 复制 Key 到环境变量

### Claude (Anthropic)

1. 访问 https://console.anthropic.com/
2. 注册账号
3. 获取 API Key

## 常见问题

### Q: 报错 "未设置 API Key"

A: 需要设置对应的环境变量：
```bash
set MOONSHOT_API_KEY=your_key_here    # 或 ZHIPU_API_KEY, ANTHROPIC_API_KEY
```

### Q: 国内无法访问 Claude API

A: 使用国内AI提供商：
```bash
python src/convert.py diagram.png --provider kimi
# 或
python src/convert.py diagram.png --provider glm
```

### Q: AI生成的SVG文字识别不准确

A: 建议：
1. 确保原图文字清晰
2. 使用高分辨率图片
3. 尝试不同AI提供商（Claude/Kimi/GLM质量有差异）

### Q: 转换后的PPT文字显示异常

A: 可能是字体问题。建议在PowerPoint中手动调整字体为系统中文字体（如微软雅黑）。

### Q: 如何查看使用的是哪个AI提供商？

A: 转换时会显示：
```
[PNG→SVG] 转换: diagram.png
  使用模型: Kimi (moonshot-v1-32k-vision-preview)
```

## 成本对比

| 提供商 | 价格/千张 | 国内访问 | 推荐场景 |
|--------|-----------|----------|----------|
| Kimi | ~¥30-50 | ✅ | 国内用户日常使用 |
| GLM | ~¥40-60 | ✅ | 中文内容优化 |
| Claude | ~$30-50 | ❌ | 高质量需求 |
| OpenAI | ~$60-100 | ❌ | 复杂图表 |

## 文件结构

```
pic2shape/
├── src/
│   ├── convert.py           # 主命令行工具
│   ├── pic2shape.py         # 批量转换工具
│   ├── pipeline.py          # PNG→SVG→PPTX完整流程
│   ├── svg_converter/       # SVG→PPTX核心模块
│   └── png2svg/             # PNG→SVG AI模块
│       ├── ai_client.py     # 多AI提供商支持
│       └── validator.py     # SVG验证器
├── CLAUDE.md                # 完整文档
└── QUICKSTART.md            # 本文件
```

## 下一步计划

- [ ] 支持更多AI提供商
- [ ] 添加颜色校正功能
- [ ] 支持批量图片预处理
- [ ] 添加转换结果缓存
