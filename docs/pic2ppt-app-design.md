# Pic2PPT 桌面应用设计方案

## 1. 项目概述

### 1.1 应用信息
- **名称**: pic2ppt
- **版本**: 1.0.0
- **定位**: AI驱动的图片转PPT可编辑形状桌面工具
- **目标平台**: Windows (主要), 后续可扩展macOS/Linux

### 1.2 核心功能
```
┌─────────────────────────────────────────────────────────────┐
│  用户上传图片  →  AI Vision生成SVG  →  转换为PPT可编辑形状  │
└─────────────────────────────────────────────────────────────┘
```

### 1.3 技术栈选择

| 组件 | 技术选择 | 理由 |
|------|----------|------|
| GUI框架 | PyQt6 | 功能强大、跨平台、成熟稳定 |
| 打包工具 | PyInstaller | 单文件EXE、支持资源嵌入 |
| 配置存储 | QSettings (Windows Registry) | 系统集成、自动持久化 |
| 日志系统 | Python logging | 标准库、灵活配置 |
| AI客户端 | OpenAI SDK | 兼容多种API格式 |
| PPT生成 | python-pptx | 成熟的PPTX操作库 |
| SVG处理 | lxml | 高性能XML解析 |

## 2. 应用架构设计

### 2.1 模块结构

```
pic2ppt/
├── pic2ppt.py              # 主入口
├── pic2ppt.spec            # PyInstaller配置
├── ico.png                 # 应用图标（占位）
├── requirements.txt        # 依赖列表
│
├── gui/                    # GUI层
│   ├── __init__.py
│   ├── main_window.py      # 主窗口
│   ├── widgets/            # 自定义组件
│   │   ├── __init__.py
│   │   ├── image_preview.py    # 图片预览区
│   │   ├── api_config.py       # API配置面板
│   │   ├── progress_panel.py   # 进度显示
│   │   └── log_viewer.py       # 日志查看器
│   └── dialogs/            # 对话框
│       ├── __init__.py
│       └── settings_dialog.py  # 设置对话框
│
├── core/                   # 业务核心层
│   ├── __init__.py
│   ├── converter.py        # 转换器封装
│   ├── ai_client.py        # AI客户端封装
│   └── config.py           # 配置管理
│
└── assets/                 # 资源文件
    └── ico.png
```

### 2.2 核心类设计

#### MainWindow (主窗口)
```python
class MainWindow(QMainWindow):
    """应用主窗口"""
    - __init__(): 初始化UI
    - _setup_ui(): 构建界面
    - _on_browse_image(): 选择图片
    - _on_convert(): 开始转换
    - _on_test_connection(): 测试AI连接
    - _update_preview(path): 更新预览
    - _show_log(): 显示日志
```

#### APIConfigWidget (API配置组件)
```python
class APIConfigWidget(QWidget):
    """AI API配置面板"""
    - provider_combo: QComboBox (提供商选择)
    - base_url_input: QLineEdit (Base URL)
    - api_key_input: QLineEdit (API Key, EchoMode=Password)
    - model_input: QLineEdit (模型名称)
    - test_btn: QPushButton (连接测试)
    - _toggle_password_visibility(): 切换密码显示
    - _test_connection(): 测试连接
```

#### ImagePreviewWidget (图片预览)
```python
class ImagePreviewWidget(QWidget):
    """图片预览组件"""
    - _load_image(path): 加载并显示图片
    - _on_drop(event): 支持拖拽
    - _zoom_in/out(): 缩放控制
```

#### ConverterThread (转换线程)
```python
class ConverterThread(QThread):
    """后台转换线程"""
    - progress_signal: 进度更新
    - log_signal: 日志消息
    - finished_signal: 完成信号
    - run(): 执行转换
```

## 3. UI界面设计

### 3.1 主窗口布局

```
┌────────────────────────────────────────────────────────────┐
│  Pic2PPT v1.0.0                                [设置][日志] │
├────────────────────────────────────────────────────────────┤
│                                                            │
│  ┌────────────────────────┐  ┌─────────────────────────┐  │
│  │                        │  │  AI API 配置            │  │
│  │                        │  │  ┌───────────────────┐  │  │
│  │                        │  │  │ 提供商: [下拉▼  ] │  │  │
│  │      图片预览区        │  │  └───────────────────┘  │  │
│  │                        │  │  ┌───────────────────┐  │  │
│  │   (支持拖拽上传)       │  │  │ Base URL: [      ]│  │  │
│  │                        │  │  └───────────────────┘  │  │
│  │   [点击或拖拽上传图片]  │  │  ┌───────────────────┐  │  │
│  │                        │  │  │ API Key: [****  ] │  │  │
│  │                        │  │  │      [显示] [连接]│  │  │
│  │                        │  │  └───────────────────┘  │  │
│  │                        │  │  ┌───────────────────┐  │  │
│  │                        │  │  │ 模型: [kimi-k2.5] │  │  │
│  └────────────────────────┘  └─────────────────────────┘  │
│                                                            │
│  ┌──────────────────────────────────────────────────────┐  │
│  │  PPT 保存位置                                        │  │
│  │  ┌──────────────────────────────┐ ┌────────────┐    │  │
│  │  │ output.pptx                  │ │ [浏览...]  │    │  │
│  │  └──────────────────────────────┘ └────────────┘    │  │
│  └──────────────────────────────────────────────────────┘  │
│                                                            │
│  ┌──────────────────────────────────────────────────────┐  │
│  │  转换进度                                            │  │
│  │  [████████████████████████████░░░░░░░░] 75%         │  │
│  │  状态: 正在生成SVG...                                │  │
│  └──────────────────────────────────────────────────────┘  │
│                                                            │
│       ┌────────────────┐    ┌────────────────┐            │
│       │   [开始转换]   │    │   [取消]       │            │
│       └────────────────┘    └────────────────┘            │
│                                                            │
└────────────────────────────────────────────────────────────┘
```

### 3.2 设置对话框

```
┌────────────────────────────────────────┐
│  设置                              [X] │
├────────────────────────────────────────┤
│  [常规] [AI模型] [输出] [关于]         │
│                                        │
│  ─── 常规设置 ───                      │
│                                        │
│  ☑ 启动时记住上次配置                  │
│                                        │
│  PPT默认尺寸:                          │
│  [宽屏16:9 ▼]                          │
│                                        │
│  默认边距: 0.5 英寸                    │
│  [====●=====]                          │
│                                        │
│  [保存]              [恢复默认]        │
└────────────────────────────────────────┘
```

## 4. 功能详细设计

### 4.1 图片上传与预览

**功能描述:**
- 支持点击选择图片文件
- 支持拖拽图片到预览区
- 支持格式: PNG, JPG, JPEG, WebP, GIF, BMP
- 图片缩略图显示，保持比例
- 显示图片信息（尺寸、大小）

**实现要点:**
```python
# 文件选择过滤器
file_filter = "图片文件 (*.png *.jpg *.jpeg *.webp *.gif *.bmp);;所有文件 (*.*)"

# 拖拽事件处理
def dragEnterEvent(event):
    if event.mimeData().hasUrls():
        event.acceptProposedAction()

def dropEvent(event):
    files = [u.toLocalFile() for u in event.mimeData().urls()]
    if files:
        self._load_image(files[0])
```

### 4.2 AI模型配置

**功能描述:**
- 提供商: 自定义（手工输入）
- Base URL: 手工输入，带默认值提示
- API Key: 密码输入框，支持显示/隐藏切换
- 模型名称: 手工输入
- 连接测试按钮：验证API连接

**连接测试逻辑:**
```python
def _test_connection():
    try:
        client = OpenAI(
            api_key=api_key,
            base_url=base_url
        )
        # 发送简单请求测试
        response = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": "hi"}],
            max_tokens=5
        )
        QMessageBox.information(self, "连接成功", "API连接正常！")
    except Exception as e:
        QMessageBox.critical(self, "连接失败", f"错误: {str(e)}")
```

### 4.3 转换流程

**流程图:**
```
开始
  │
  ▼
检查输入 ──→ 输入无效 → 提示错误
  │
  ▼
读取图片
  │
  ▼
AI生成SVG ──→ 失败 → 重试/取消
  │
  ▼
验证SVG
  │
  ▼
转换PPT ──→ 失败 → 错误日志
  │
  ▼
保存PPT
  │
  ▼
完成
```

**进度显示:**
- 0-30%: 上传图片到AI
- 30-70%: AI生成SVG
- 70-90%: 转换SVG到PPT
- 90-100%: 保存文件

### 4.4 日志系统

**日志配置:**
```python
import logging
from pathlib import Path

log_path = Path(__file__).parent / "pic2ppt.log"

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_path, encoding='utf-8'),
        logging.StreamHandler()
    ]
)
```

**日志内容:**
- 应用启动/关闭
- 文件操作（读取、保存）
- AI API调用（请求、响应时间、错误）
- 转换进度
- 错误堆栈

**日志查看器:**
- 实时滚动显示
- 支持过滤（INFO/WARNING/ERROR）
- 支持清空/导出

## 5. 配置持久化

### 5.1 存储项

| 配置项 | 类型 | 默认值 |
|--------|------|--------|
| provider | str | "" |
| base_url | str | "" |
| model | str | "" |
| save_dir | str | 用户文档目录 |
| slide_ratio | str | "16:9" |
| margin | float | 0.5 |
| remember_config | bool | True |
| window_geometry | QByteArray | - |

### 5.2 安全考虑

- API Key使用QSettings加密存储（Windows DPAPI）
- 提供"清除所有配置"选项

## 6. 打包部署

### 6.1 PyInstaller配置

**pic2ppt.spec:**
```python
# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['pic2ppt.py'],
    pathex=['.'],
    binaries=[],
    datas=[('src', 'src'), ('assets', 'assets')],
    hiddenimports=['openai', 'pptx', 'lxml', 'PyQt6'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='pic2ppt',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI模式
    icon='ico.png',  # 应用图标
)
```

### 6.2 构建命令

```bash
# 安装依赖
pip install -r requirements.txt

# 打包
pyinstaller pic2ppt.spec --clean

# 输出
# dist/pic2ppt.exe
```

### 6.3 依赖列表

```
PyQt6>=6.4.0
openai>=1.0.0
python-pptx>=0.6.21
lxml>=4.9.0
Pillow>=9.0.0
requests>=2.28.0
```

## 7. 异常处理设计

### 7.1 错误分类

| 错误类型 | 场景 | 处理方式 |
|----------|------|----------|
| 输入错误 | 未选择图片 | 弹窗提示 |
| 配置错误 | API Key为空 | 弹窗提示 |
| 网络错误 | API连接失败 | 重试3次，失败后提示 |
| AI错误 | 生成失败 | 记录日志，提示用户 |
| 转换错误 | SVG解析失败 | 尝试修复，失败提示 |
| 文件错误 | 保存失败 | 提示选择其他路径 |

### 7.2 用户友好提示

```python
ERROR_MESSAGES = {
    "NO_IMAGE": "请先选择要转换的图片",
    "NO_API_KEY": "请配置API Key",
    "API_ERROR": "AI服务暂时不可用，请稍后重试",
    "CONVERT_ERROR": "转换失败，请检查图片内容",
    "SAVE_ERROR": "无法保存到指定位置，请检查权限",
}
```

## 8. 性能优化

### 8.1 异步处理

- AI调用在独立线程执行，不阻塞UI
- 文件IO使用异步操作
- 进度更新通过信号槽机制

### 8.2 内存管理

- 大图预览使用缩略图
- 及时释放临时SVG文件
- 单文件处理，不缓存过多数据

## 9. 后续扩展建议

### 9.1 功能扩展

- [ ] 批量转换
- [ ] SVG编辑功能
- [ ] 模板系统
- [ ] 历史记录
- [ ] 多语言支持

### 9.2 技术扩展

- [ ] 支持更多AI提供商
- [ ] 本地模型支持
- [ ] 插件系统

## 10. 项目时间线

| 阶段 | 内容 | 预计时间 |
|------|------|----------|
| 1 | 核心框架搭建 | 1天 |
| 2 | GUI界面实现 | 2天 |
| 3 | 转换流程集成 | 1天 |
| 4 | 打包测试 | 1天 |
| **总计** | | **5天** |

---

文档版本: 1.0
最后更新: 2026-03-16
