#!/usr/bin/env python3
"""pic2ppt - AI图片转PPT可编辑形状

绿色便携版，单文件EXE，双击即运行
支持三种上传方式：文件选择、拖拽、粘贴
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
import json
import logging
import sys
import os
import tempfile
import threading
import base64

# 使用限制
from src.usage_limiter import UsageLimiter

# PyInstaller 隐藏导入 - 确保这些模块被打包
# anthropic 和 openai 用于 API 连接
from PIL import Image, ImageTk
import anthropic
import openai
import httpx

# 尝试导入 tkinterdnd2 用于拖拽支持
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    TKDND_AVAILABLE = True
except ImportError:
    TKDND_AVAILABLE = False

# 添加src到路径
if getattr(sys, 'frozen', False):
    # 打包后的环境
    BASE_DIR = Path(sys._MEIPASS)
else:
    # 开发环境
    BASE_DIR = Path(__file__).parent

sys.path.insert(0, str(BASE_DIR / "src"))

# 配置日志
log_file = Path(__file__).parent / "pic2ppt.log"
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)
logger = logging.getLogger('pic2ppt')


class CollapsibleFrame(ttk.Frame):
    """可折叠面板"""
    def __init__(self, parent, title="", *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.is_expanded = True

        # 标题栏
        self.header = ttk.Frame(self)
        self.header.pack(fill=tk.X, pady=2)

        self.toggle_btn = ttk.Button(
            self.header, text="▼", width=3,
            command=self._toggle
        )
        self.toggle_btn.pack(side=tk.LEFT, padx=2)

        self.title_label = ttk.Label(self.header, text=title, font=('Microsoft YaHei', 10, 'bold'))
        self.title_label.pack(side=tk.LEFT, padx=5)

        # 内容区
        self.content = ttk.Frame(self)
        self.content.pack(fill=tk.X, expand=True)

    def _toggle(self):
        self.is_expanded = not self.is_expanded
        if self.is_expanded:
            self.toggle_btn.config(text="▼")
            self.content.pack(fill=tk.X, expand=True)
        else:
            self.toggle_btn.config(text="▶")
            self.content.pack_forget()


class Pic2PPTApp:
    def __init__(self, root):
        self.root = root
        self.root.title("pic2ppt v1.0")
        self.root.geometry("620x780")
        self.root.minsize(620, 780)
        self.root.resizable(False, False)

        # 初始化使用限制
        self.limiter = UsageLimiter()
        
        # 检查单实例
        ok, msg = self.limiter.check_single_instance()
        if not ok:
            messagebox.showerror("启动失败", msg)
            sys.exit(1)
        
        # 检查每日限制
        ok, msg = self.limiter.check_daily_limit()
        if not ok:
            messagebox.showerror("使用限制", msg)
            sys.exit(1)
        
        # 当前图片路径
        self.current_image_path = None
        self.is_converting = False
        self.convert_thread = None

        # 加载配置
        self.config = self._load_config()

        # 设置样式
        self._setup_styles()

        # 构建界面
        self._setup_ui()

        # 绑定事件
        self._bind_events()

        # 启用文件拖拽（Windows）
        self._enable_drag_drop()

        logger.info("应用启动成功")

    def _setup_styles(self):
        """设置样式"""
        style = ttk.Style()
        style.theme_use('clam')

        # 自定义样式
        style.configure('Title.TLabel', font=('Microsoft YaHei', 12, 'bold'))
        style.configure('Status.TLabel', font=('Microsoft YaHei', 9))
        style.configure('Action.TButton', font=('Microsoft YaHei', 10))

    def _load_header_icon(self, parent):
        """加载标题栏图标 - 直接绘制SVG风格的大图标"""
        self._draw_svg_style_icon(parent)

    def _draw_svg_style_icon(self, parent):
        """绘制SVG风格的完整图标（60x60，清晰展示图案和文字）"""
        size = 60
        icon_canvas = tk.Canvas(parent, width=size, height=size, bg='#f0f0f0',
                                highlightthickness=1, highlightbackground='#ddd',
                                bd=0)
        icon_canvas.pack(side=tk.LEFT, padx=(0, 8))

        # 颜色定义（与SVG一致）
        color_left = '#FB923C'      # 左橙渐变
        color_right = '#FBBF24'     # 右琥珀渐变
        bg_color = '#FFFBEB'        # 浅黄背景
        dark_orange = '#EA580C'     # 深色用于图案

        # 绘制背景圆角矩形（使用椭圆弧线模拟）
        self._draw_rounded_rect(icon_canvas, 2, 2, size-2, size-2, 8, fill=bg_color, outline='')

        # 方块配置
        block_size = 22
        gap = 3
        margin = 6

        # 四个方块位置（左上、右上、左下、右下）
        positions = [
            (margin, margin),                    # 左上
            (margin + block_size + gap, margin), # 右上
            (margin, margin + block_size + gap), # 左下
            (margin + block_size + gap, margin + block_size + gap), # 右下
        ]

        colors = [color_left, color_right, color_left, color_right]

        # 绘制带内部图案的方块
        for idx, ((x, y), color) in enumerate(zip(positions, colors)):
            # 方块主体（带圆角）
            self._draw_rounded_rect(icon_canvas, x, y, x + block_size, y + block_size,
                                    4, fill=color, outline='', width=0)

            # 在方块内绘制简单图案
            cx = x + block_size // 2
            cy = y + block_size // 2

            if idx % 2 == 0:  # 左侧方块 - 图片图标（简单山形）
                # 白色小矩形作为图片框
                icon_canvas.create_rectangle(
                    cx-6, cy-4, cx+6, cy+6,
                    fill='white', outline='', width=0
                )
                # 小圆点（太阳）
                icon_canvas.create_oval(
                    cx-4, cy-2, cx-1, cy+1,
                    fill=dark_orange, outline='', width=0
                )
            else:  # 右侧方块 - PPT图标（三条横线）
                line_w = 10
                line_h = 2
                for i in range(3):
                    line_y = cy - 3 + i * 4
                    line_w_i = line_w - i * 2
                    icon_canvas.create_rectangle(
                        cx - line_w_i//2, line_y,
                        cx + line_w_i//2, line_y + line_h,
                        fill='white', outline='', width=0
                    )

        # 中央箭头（旋转效果）
        arrow_cx = size // 2
        arrow_cy = size // 2
        # 白色圆形背景
        icon_canvas.create_oval(
            arrow_cx-6, arrow_cy-6, arrow_cx+6, arrow_cy+6,
            fill='white', outline='', width=0
        )
        # 简化的右箭头
        icon_canvas.create_polygon(
            arrow_cx-3, arrow_cy-2,
            arrow_cx+2, arrow_cy-2,
            arrow_cx+2, arrow_cy-4,
            arrow_cx+5, arrow_cy,
            arrow_cx+2, arrow_cy+4,
            arrow_cx+2, arrow_cy+2,
            arrow_cx-3, arrow_cy+2,
            fill='#D97706', outline='', width=0
        )

        # 底部 FangJin 文字（小字号，紧凑）
        icon_canvas.create_text(
            size // 2, size - 8,
            text="FangJin",
            font=('Arial', 6, 'bold'),
            fill='#92400E',
            anchor='center'
        )

    def _draw_rounded_rect(self, canvas, x1, y1, x2, y2, radius, fill='', outline='', width=1):
        """绘制圆角矩形"""
        # 使用圆弧和直线模拟圆角矩形
        # 左上圆弧
        canvas.create_arc(x1, y1, x1 + 2*radius, y1 + 2*radius,
                          start=90, extent=90, style='pieslice',
                          fill=fill, outline=outline, width=width)
        # 右上圆弧
        canvas.create_arc(x2 - 2*radius, y1, x2, y1 + 2*radius,
                          start=0, extent=90, style='pieslice',
                          fill=fill, outline=outline, width=width)
        # 右下圆弧
        canvas.create_arc(x2 - 2*radius, y2 - 2*radius, x2, y2,
                          start=270, extent=90, style='pieslice',
                          fill=fill, outline=outline, width=width)
        # 左下圆弧
        canvas.create_arc(x1, y2 - 2*radius, x1 + 2*radius, y2,
                          start=180, extent=90, style='pieslice',
                          fill=fill, outline=outline, width=width)
        # 中间矩形
        canvas.create_rectangle(x1 + radius, y1, x2 - radius, y2,
                                fill=fill, outline='', width=0)
        canvas.create_rectangle(x1, y1 + radius, x1 + radius, y2 - radius,
                                fill=fill, outline='', width=0)
        canvas.create_rectangle(x2 - radius, y1 + radius, x2, y2 - radius,
                                fill=fill, outline='', width=0)
        canvas.create_rectangle(x1 + radius, y1 + radius, x2 - radius, y2 - radius,
                                fill=fill, outline='', width=0)

    def _draw_simple_icon(self, parent):
        """绘制简单的图标（备用）"""
        icon_canvas = tk.Canvas(parent, width=140, height=60, bg='#f0f0f0',
                                highlightthickness=0, bd=0)
        icon_canvas.pack(side=tk.LEFT, padx=(0, 5))

        # 绘制2x2方块（左橙右黄）
        block_size = 22
        gap = 3
        colors = ['#F97316', '#F59E0B', '#F97316', '#F59E0B']
        positions = [(8, 8), (33, 8), (8, 33), (33, 33)]

        for (x, y), color in zip(positions, colors):
            icon_canvas.create_rectangle(
                x, y, x + block_size, y + block_size,
                fill=color, outline='', width=0
            )

        # 添加文字
        icon_canvas.create_text(
            100, 30,
            text="FangJin",
            font=('Arial', 11, 'bold'),
            fill='#92400E',
            anchor='center'
        )

    def _setup_ui(self):
        """构建界面"""
        # 主容器
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # === 标题 ===
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))

        # 加载并显示图标
        self._load_header_icon(header_frame)

        title_label = ttk.Label(
            header_frame,
            text="pic2ppt",
            font=('Microsoft YaHei', 16, 'bold')
        )
        title_label.pack(side=tk.LEFT, padx=(5, 0))

        version_label = ttk.Label(
            header_frame,
            text="v1.0",
            foreground="#666"
        )
        version_label.pack(side=tk.LEFT, padx=(5, 0))

        # 右上角按钮
        btn_frame = ttk.Frame(header_frame)
        btn_frame.pack(side=tk.RIGHT)

        ttk.Button(btn_frame, text="日志", command=self._show_log, width=6).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="?", command=self._show_help, width=3).pack(side=tk.LEFT, padx=2)

        # === 图片预览区 ===
        preview_frame = ttk.LabelFrame(main_frame, text=" 图片预览 ", padding="5")
        preview_frame.pack(fill=tk.X, pady=5)

        # 预览画布
        self.preview_canvas = tk.Canvas(
            preview_frame,
            width=560, height=260,
            bg="#f5f5f5",
            highlightthickness=1,
            highlightbackground="#ccc"
        )
        self.preview_canvas.pack()

        # 提示文字 - 根据拖拽支持动态显示
        drag_hint = "\n或拖拽文件" if TKDND_AVAILABLE else ""
        self.preview_text = self.preview_canvas.create_text(
            280, 130,
            text=f"点击选择图片{drag_hint}\n或 Ctrl+V 粘贴",
            font=('Microsoft YaHei', 11),
            fill="#666",
            justify=tk.CENTER
        )

        # 图片标签（初始隐藏）
        self.preview_image = None
        self.preview_image_id = None

        # 图片信息
        self.image_info_label = ttk.Label(preview_frame, text="", foreground="#666")
        self.image_info_label.pack(pady=(5, 0))

        # === API配置区（可折叠） ===
        self.api_frame = CollapsibleFrame(main_frame, title="AI API 配置")
        self.api_frame.pack(fill=tk.X, pady=5)

        # Base URL
        ttk.Label(self.api_frame.content, text="Base URL:").grid(row=0, column=0, sticky=tk.W, pady=3)
        self.base_url_var = tk.StringVar(value=self.config.get('base_url', ''))
        self.base_url_entry = ttk.Entry(
            self.api_frame.content,
            textvariable=self.base_url_var,
            width=55
        )
        self.base_url_entry.grid(row=0, column=1, columnspan=2, sticky=tk.EW, pady=3)

        # API Key
        ttk.Label(self.api_frame.content, text="API Key:").grid(row=1, column=0, sticky=tk.W, pady=3)
        self.api_key_var = tk.StringVar(value=self.config.get('api_key', ''))
        self.api_key_entry = ttk.Entry(
            self.api_frame.content,
            textvariable=self.api_key_var,
            width=45,
            show="*"
        )
        self.api_key_entry.grid(row=1, column=1, sticky=tk.EW, pady=3)

        # 显示/隐藏按钮
        self.show_key_btn = ttk.Button(
            self.api_frame.content,
            text="👁",
            width=3,
            command=self._toggle_key_visibility
        )
        self.show_key_btn.grid(row=1, column=2, padx=2)

        # 模型
        ttk.Label(self.api_frame.content, text="模型:").grid(row=2, column=0, sticky=tk.W, pady=3)
        self.model_var = tk.StringVar(value=self.config.get('model', 'kimi-k2.5'))
        self.model_entry = ttk.Entry(
            self.api_frame.content,
            textvariable=self.model_var,
            width=40
        )
        self.model_entry.grid(row=2, column=1, sticky=tk.W, pady=3)

        # 测试连接按钮
        self.test_btn = ttk.Button(
            self.api_frame.content,
            text="测试连接",
            command=self._test_connection,
            width=10
        )
        self.test_btn.grid(row=2, column=2, padx=5)

        # 连接状态
        self.connection_status = ttk.Label(self.api_frame.content, text="", foreground="#666")
        self.connection_status.grid(row=3, column=1, columnspan=2, sticky=tk.W)

        self.api_frame.content.columnconfigure(1, weight=1)

        # === 输出设置（可折叠） ===
        self.output_frame = CollapsibleFrame(main_frame, title="输出设置")
        self.output_frame.pack(fill=tk.X, pady=5)

        # 保存路径
        path_frame = ttk.Frame(self.output_frame.content)
        path_frame.pack(fill=tk.X, pady=3)

        ttk.Label(path_frame, text="保存位置:").pack(side=tk.LEFT)

        self.output_path_var = tk.StringVar(value=self.config.get('output_dir', str(Path.home() / "Documents")))
        self.output_entry = ttk.Entry(path_frame, textvariable=self.output_path_var, width=40)
        self.output_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        ttk.Button(path_frame, text="浏览...", command=self._browse_output, width=8).pack(side=tk.LEFT)

        # === 进度区域 ===
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=10)

        # 进度条
        self.progress = ttk.Progressbar(
            progress_frame,
            length=560,
            mode='determinate'
        )
        self.progress.pack(fill=tk.X)

        # 状态标签 - 显示剩余次数
        remaining = self.limiter.get_remaining()
        self.status_label = ttk.Label(
            progress_frame,
            text=f"就绪 - 请选择图片（今日剩余 {remaining} 次）",
            foreground="#333"
        )
        self.status_label.pack(pady=(5, 0))

        # === 按钮区域 ===
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)

        self.convert_btn = ttk.Button(
            button_frame,
            text="开始转换",
            command=self._on_convert,
            width=15
        )
        self.convert_btn.pack(side=tk.LEFT, padx=10)

        self.cancel_btn = ttk.Button(
            button_frame,
            text="取消",
            command=self._on_cancel,
            width=15,
            state=tk.DISABLED
        )
        self.cancel_btn.pack(side=tk.LEFT, padx=10)

    def _bind_events(self):
        """绑定事件"""
        # 点击预览区
        self.preview_canvas.bind('<Button-1>', lambda e: self._browse_image())

        # 粘贴快捷键
        self.root.bind('<Control-v>', self._on_paste)
        self.root.bind('<Control-V>', self._on_paste)

        # 窗口关闭
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _enable_drag_drop(self):
        """启用文件拖拽支持"""
        if not TKDND_AVAILABLE:
            logger.info("tkinterdnd2 未安装，拖拽功能不可用")
            return

        try:
            # 注册预览区域为拖拽目标
            self.preview_canvas.drop_target_register(DND_FILES)
            self.preview_canvas.dnd_bind('<<Drop>>', self._on_canvas_drop)
            logger.info("文件拖拽功能已启用")
        except Exception as e:
            logger.error(f"启用拖拽失败: {e}")

    def _on_canvas_drop(self, event):
        """处理拖拽事件"""
        try:
            # tkinterdnd2 返回的文件路径可能包含花括号，需要处理
            file_path = event.data.strip('{}')
            logger.info(f"拖拽文件: {file_path}")

            # 检查是否是支持的图片格式
            supported_exts = ('.png', '.jpg', '.jpeg', '.webp', '.gif', '.bmp', '.tiff', '.tif')
            if file_path.lower().endswith(supported_exts):
                self._load_image(file_path)
            else:
                logger.warning(f"不支持的文件格式: {file_path}")
                messagebox.showwarning("格式不支持", "请拖拽 PNG、JPG、JPEG 等图片文件")
        except Exception as e:
            logger.error(f"处理拖拽文件失败: {e}")

    def _on_drop_file(self, file_path):
        """处理拖拽的文件"""
        try:
            logger.info(f"处理拖拽文件: {file_path}")
            if os.path.exists(file_path):
                self._load_image(file_path)
            else:
                logger.error(f"文件不存在: {file_path}")
        except Exception as e:
            logger.error(f"加载拖拽文件失败: {e}")
            messagebox.showerror("错误", f"无法加载图片:\n{str(e)}")

    def _browse_image(self):
        """浏览选择图片"""
        filetypes = [
            ("图片文件", "*.png *.jpg *.jpeg *.webp *.gif *.bmp"),
            ("PNG", "*.png"),
            ("JPEG", "*.jpg *.jpeg"),
            ("WebP", "*.webp"),
            ("所有文件", "*.*")
        ]

        path = filedialog.askopenfilename(
            title="选择图片",
            filetypes=filetypes
        )

        if path:
            self._load_image(path)

    def _load_image(self, path):
        """加载并显示图片"""
        try:
            from PIL import Image, ImageTk

            self.current_image_path = path

            # 打开图片
            img = Image.open(path)
            orig_width, orig_height = img.size

            # 计算缩放比例（适应预览区 560x280）
            scale_w = 540 / orig_width
            scale_h = 260 / orig_height
            scale = min(scale_w, scale_h, 1.0)

            new_width = int(orig_width * scale)
            new_height = int(orig_height * scale)

            # 缩放图片
            img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)

            # 转换为PhotoImage
            self.preview_image = ImageTk.PhotoImage(img)

            # 清除提示文字
            self.preview_canvas.delete(self.preview_text)

            # 显示图片（居中）
            if self.preview_image_id:
                self.preview_canvas.delete(self.preview_image_id)

            x = (560 - new_width) // 2
            y = (280 - new_height) // 2

            self.preview_image_id = self.preview_canvas.create_image(
                x, y, anchor=tk.NW, image=self.preview_image
            )

            # 更新信息
            file_size = Path(path).stat().st_size / 1024
            self.image_info_label.config(
                text=f"{orig_width}x{orig_height} | {file_size:.1f} KB | {Path(path).name[:30]}"
            )

            # 自动设置输出文件名
            output_dir = Path(self.output_path_var.get())
            output_name = Path(path).stem + ".pptx"
            self.output_path_var.set(str(output_dir / output_name))

            self._update_status("图片已加载，请配置API后点击转换")
            logger.info(f"加载图片: {path}")

        except Exception as e:
            logger.error(f"加载图片失败: {e}")
            messagebox.showerror("错误", f"无法加载图片:\n{str(e)}")

    def _on_paste(self, event=None):
        """从剪贴板粘贴 - 只在预览画布有焦点时处理"""
        # 检查当前焦点是否在预览画布
        try:
            focused_widget = self.root.focus_get()
            if focused_widget != self.preview_canvas:
                # 如果焦点在其他输入框，让系统正常处理粘贴
                return
        except:
            pass

        try:
            from PIL import Image, ImageGrab

            # 获取剪贴板图片
            img = ImageGrab.grabclipboard()

            if isinstance(img, Image.Image):
                # 保存临时文件
                temp_path = Path(tempfile.gettempdir()) / "pic2ppt_paste.png"
                img.save(temp_path)
                self._load_image(str(temp_path))
                self._update_status("已从剪贴板粘贴图片")
            else:
                messagebox.showinfo("提示", "剪贴板中没有图片")

        except Exception as e:
            logger.error(f"粘贴失败: {e}")
            messagebox.showerror("错误", f"粘贴失败:\n{str(e)}")

    def _toggle_key_visibility(self):
        """切换API Key显示/隐藏"""
        if self.api_key_entry.cget('show') == '*':
            self.api_key_entry.config(show='')
            self.show_key_btn.config(text='🙈')
        else:
            self.api_key_entry.config(show='*')
            self.show_key_btn.config(text='👁')

    def _test_connection(self):
        """测试AI连接"""
        base_url = self.base_url_var.get().strip()
        api_key = self.api_key_var.get().strip()
        model = self.model_var.get().strip()

        if not api_key:
            messagebox.showwarning("提示", "请输入API Key")
            return

        self.test_btn.config(state=tk.DISABLED, text="测试中...")
        self.connection_status.config(text="正在测试连接...", foreground="#666")

        def test():
            logger.info(f"开始测试连接: URL={base_url}, Model={model}")
            error_messages = []

            # 方法1: 使用Anthropic SDK（优先，因为配置的可能是Anthropic端点）
            try:
                logger.info("尝试使用Anthropic SDK...")
                import anthropic

                client = anthropic.Anthropic(
                    api_key=api_key,
                    base_url=base_url if base_url else None
                )

                logger.info(f"Anthropic SDK客户端创建成功, base_url={base_url}")

                # Anthropic 使用 messages.create 而不是 chat.completions.create
                response = client.messages.create(
                    model=model,
                    max_tokens=5,
                    messages=[{"role": "user", "content": "hi"}]
                )

                logger.info(f"Anthropic SDK测试成功: {response}")
                self.root.after(0, lambda: self._test_success())
                return

            except ImportError:
                logger.warning("Anthropic SDK未安装，跳过")
                error_messages.append("Anthropic SDK未安装")
            except Exception as e:
                error_msg = f"Anthropic SDK: {type(e).__name__}: {str(e)[:100]}"
                logger.error(error_msg)
                error_messages.append(error_msg)

            # 方法2: 使用OpenAI SDK
            try:
                logger.info("尝试使用OpenAI SDK...")
                from openai import OpenAI

                client = OpenAI(
                    api_key=api_key,
                    base_url=base_url if base_url else None
                )

                logger.info(f"OpenAI SDK客户端创建成功, base_url={base_url}")
                response = client.chat.completions.create(
                    model=model,
                    messages=[{"role": "user", "content": "hi"}],
                    max_tokens=5,
                    timeout=15
                )

                logger.info(f"OpenAI SDK测试成功: {response}")
                self.root.after(0, lambda: self._test_success())
                return

            except ImportError:
                logger.warning("OpenAI SDK未安装，跳过")
                error_messages.append("OpenAI SDK未安装")
            except Exception as e:
                error_msg = f"OpenAI SDK: {type(e).__name__}: {str(e)[:100]}"
                logger.error(error_msg)
                error_messages.append(error_msg)

            # 方法3: 使用urllib作为备选
            try:
                logger.info("尝试使用urllib...")
                import urllib.request
                import urllib.error
                import json
                import ssl

                # 构建请求URL
                if base_url:
                    url = base_url.rstrip('/') + '/chat/completions'
                else:
                    url = 'https://api.openai.com/v1/chat/completions'

                logger.info(f"测试URL: {url}")

                # 准备请求数据
                data = {
                    "model": model,
                    "messages": [{"role": "user", "content": "hi"}],
                    "max_tokens": 5
                }

                # 创建请求
                req = urllib.request.Request(
                    url,
                    data=json.dumps(data).encode('utf-8'),
                    headers={
                        'Content-Type': 'application/json',
                        'Authorization': f'Bearer {api_key}'
                    },
                    method='POST'
                )

                # 禁用SSL证书验证
                ctx = ssl.create_default_context()
                ctx.check_hostname = False
                ctx.verify_mode = ssl.CERT_NONE

                # 设置超时
                logger.info("发送HTTP请求...")
                response = urllib.request.urlopen(req, timeout=15, context=ctx)

                if response.status == 200:
                    logger.info(f"HTTP响应成功: {response.status}")
                    self.root.after(0, lambda: self._test_success())
                    return
                else:
                    logger.warning(f"HTTP响应异常: {response.status}")
                    error_messages.append(f"HTTP {response.status}")

            except urllib.error.HTTPError as e:
                error_body = e.read().decode('utf-8', errors='ignore')
                error_msg = f"HTTP {e.code}: {error_body[:200]}"
                logger.error(f"HTTP错误: {error_msg}")
                error_messages.append(error_msg)
            except urllib.error.URLError as e:
                error_msg = f"网络错误: {e.reason}"
                logger.error(error_msg)
                error_messages.append(error_msg)
            except Exception as e:
                error_msg = f"urllib: {type(e).__name__}: {str(e)[:100]}"
                logger.error(error_msg)
                error_messages.append(error_msg)

            # 所有方法都失败
            final_error = error_messages[0] if error_messages else "未知错误"
            logger.error(f"所有测试方法都失败: {error_messages}")
            self.root.after(0, lambda err=final_error: self._test_failed(err))

        threading.Thread(target=test, daemon=True).start()

    def _test_success(self):
        """连接测试成功"""
        self.test_btn.config(state=tk.NORMAL, text="测试连接")
        self.connection_status.config(text="✓ 连接正常", foreground="#26a269")
        logger.info("API连接测试成功")

    def _test_failed(self, error):
        """连接测试失败"""
        self.test_btn.config(state=tk.NORMAL, text="测试连接")
        self.connection_status.config(text=f"✗ 连接失败: {error[:50]}", foreground="#c00")
        logger.error(f"API连接测试失败: {error}")

    def _browse_output(self):
        """选择输出目录"""
        directory = filedialog.askdirectory(
            title="选择保存位置",
            initialdir=self.output_path_var.get()
        )
        if directory:
            if self.current_image_path:
                output_name = Path(self.current_image_path).stem + ".pptx"
                self.output_path_var.set(str(Path(directory) / output_name))
            else:
                self.output_path_var.set(directory)

    def _on_convert(self):
        """开始转换"""
        if not self.current_image_path:
            messagebox.showwarning("提示", "请先选择图片")
            return

        # 检查API配置
        api_key = self.api_key_var.get().strip()
        if not api_key:
            messagebox.showwarning("提示", "请输入API Key")
            return

        # 保存配置
        self._save_config()

        # 检查每日限制
        remaining = self.limiter.get_remaining()
        if remaining <= 0:
            messagebox.showwarning("使用限制", "今日转换次数已达上限（3次），请明天再试")
            return

        # 更新UI状态
        self.is_converting = True
        self.convert_btn.config(state=tk.DISABLED)
        self.cancel_btn.config(state=tk.NORMAL)
        self.progress['value'] = 0

        # 获取参数
        base_url = self.base_url_var.get().strip()
        model = self.model_var.get().strip() or "kimi-k2.5"
        output_path = self.output_path_var.get()

        # 启动转换线程
        self.convert_thread = threading.Thread(
            target=self._do_convert,
            args=(self.current_image_path, output_path, base_url, api_key, model),
            daemon=True
        )
        self.convert_thread.start()

    def _do_convert(self, image_path, output_path, base_url, api_key, model):
        """执行转换（后台线程）"""
        try:
            self._update_progress(10, "正在读取图片...")
            logger.info(f"开始转换: {image_path}")

            # 导入转换模块
            from src.pipeline import PNGToPPTXPipeline

            self._update_progress(20, "初始化AI客户端...")

            # 创建pipeline - 根据base_url智能选择provider
            # 如果base_url包含anthropic，使用claude provider，否则使用openai
            provider = "claude" if base_url and "anthropic" in base_url.lower() else "openai"
            logger.info(f"使用provider: {provider}, model: {model}, base_url: {base_url}")

            pipeline = PNGToPPTXPipeline(
                provider=provider,
                api_key=api_key,
                base_url=base_url,
                model=model
            )

            self._update_progress(30, "AI生成SVG中，请稍候...")

            # 转换
            def progress_callback(step, total, message):
                progress = 30 + (step / total) * 60
                self._update_progress(progress, message)

            pipeline.set_progress_callback(progress_callback)

            result = pipeline.convert(
                image_path=image_path,
                output_pptx=output_path,
                keep_svg=False
            )

            self._update_progress(100, "转换完成！")
            self.root.after(0, lambda: self._convert_success(output_path))

        except Exception as e:
            logger.exception("转换失败")
            self.root.after(0, lambda: self._convert_failed(str(e)))

    def _update_progress(self, value, message):
        """更新进度（线程安全）"""
        self.root.after(0, lambda: self._do_update_progress(value, message))

    def _do_update_progress(self, value, message):
        """实际更新进度"""
        self.progress['value'] = value
        self._update_status(message)
        self.root.update_idletasks()

    def _update_status(self, message):
        """更新状态文字"""
        self.status_label.config(text=message)

    def _convert_success(self, output_path):
        """转换成功"""
        self.is_converting = False
        self.convert_btn.config(state=tk.NORMAL)
        self.cancel_btn.config(state=tk.DISABLED)

        # 记录转换成功
        self._record_conversion()

        # 更新状态显示剩余次数
        remaining = self.limiter.get_remaining()
        self.status_label.config(text=f"转换完成！今日剩余 {remaining} 次")

        # 询问是否打开文件
        if messagebox.askyesno(
            "转换成功",
            f"PPT已保存到:\n{output_path}\n\n是否打开文件？"
        ):
            import os
            os.startfile(output_path)

        logger.info(f"转换成功: {output_path}")

    def _record_conversion(self):
        """记录转换成功"""
        if hasattr(self, 'limiter'):
            self.limiter.record_conversion()

    def _convert_failed(self, error):
        """转换失败"""
        self.is_converting = False
        self.convert_btn.config(state=tk.NORMAL)
        self.cancel_btn.config(state=tk.DISABLED)
        self.progress['value'] = 0
        self._update_status("转换失败")

        messagebox.showerror("转换失败", f"错误信息:\n{error}")
        logger.error(f"转换失败: {error}")

    def _on_cancel(self):
        """取消转换"""
        # 设置标志让线程退出
        self.is_converting = False
        self._update_status("已取消")
        self.convert_btn.config(state=tk.NORMAL)
        self.cancel_btn.config(state=tk.DISABLED)
        self.progress['value'] = 0

    def _load_config(self):
        """加载配置"""
        config_file = Path(__file__).parent / "config.json"
        if config_file.exists():
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                logger.error(f"加载配置失败: {e}")

        return {
            'base_url': '',
            'api_key': '',
            'model': 'kimi-k2.5',
            'output_dir': str(Path.home() / "Documents")
        }

    def _save_config(self):
        """保存配置"""
        config = {
            'base_url': self.base_url_var.get().strip(),
            'api_key': self.api_key_var.get().strip(),
            'model': self.model_var.get().strip(),
            'output_dir': str(Path(self.output_path_var.get()).parent)
        }

        try:
            config_file = Path(__file__).parent / "config.json"
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            logger.info("配置已保存")
        except Exception as e:
            logger.error(f"保存配置失败: {e}")

    def _show_log(self):
        """显示日志窗口"""
        log_window = tk.Toplevel(self.root)
        log_window.title("运行日志")
        log_window.geometry("600x400")

        # 日志文本框
        text = scrolledtext.ScrolledText(log_window, wrap=tk.WORD)
        text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 读取日志
        try:
            if log_file.exists():
                with open(log_file, 'r', encoding='utf-8') as f:
                    text.insert(tk.END, f.read())
            else:
                text.insert(tk.END, "暂无日志")
        except Exception as e:
            text.insert(tk.END, f"读取日志失败: {e}")

        text.config(state=tk.DISABLED)

        # 按钮
        ttk.Button(log_window, text="清空日志", command=self._clear_log).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(log_window, text="关闭", command=log_window.destroy).pack(side=tk.RIGHT, padx=5, pady=5)

    def _clear_log(self):
        """清空日志"""
        try:
            with open(log_file, 'w', encoding='utf-8') as f:
                f.write("")
            messagebox.showinfo("提示", "日志已清空")
        except Exception as e:
            messagebox.showerror("错误", f"清空日志失败: {e}")

    def _show_help(self):
        """显示帮助"""
        drag_help = "、拖拽文件" if TKDND_AVAILABLE else ""
        help_text = f"""pic2ppt - AI图片转PPT可编辑形状

使用步骤:
1. 上传图片：点击预览区选择{drag_help}、或 Ctrl+V 粘贴
2. 配置API：输入Base URL、API Key和模型名称
3. 测试连接：点击"测试连接"验证API可用
4. 设置输出：选择PPT保存位置
5. 开始转换：点击"开始转换"按钮

支持格式: PNG, JPG, JPEG, WebP, GIF, BMP

问题反馈: 查看 pic2ppt.log 获取详细日志
"""
        messagebox.showinfo("帮助", help_text)

    def _on_close(self):
        """关闭窗口"""
        if self.is_converting:
            if not messagebox.askyesno("确认", "转换正在进行中，确定要退出吗？"):
                return

        # 释放实例锁
        if hasattr(self, 'limiter'):
            self.limiter.release_lock()

        self._save_config()
        self.root.destroy()


def main():
    """主入口"""
    # 使用 TkinterDnD.Tk 支持拖拽（如果可用）
    if TKDND_AVAILABLE:
        root = TkinterDnD.Tk()
        logger.info("使用 TkinterDnD 模式（支持拖拽）")
    else:
        root = tk.Tk()
        logger.info("使用标准 Tk 模式（拖拽不可用）")

    # 设置DPI感知（Windows）
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass

    app = Pic2PPTApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
