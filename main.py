import os
import sys
import ctypes
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import fitz  # PyMuPDF
import threading
import configparser
import queue
import numpy as np  # 添加numpy库
from PIL import Image  # 添加PIL库用于处理图片

# 判断是否在打包环境中运行
def resource_path(relative_path):
    """获取资源的绝对路径，兼容PyInstaller打包后的情况"""
    try:
        # PyInstaller创建临时文件夹，将路径存储在_MEIPASS中
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_config_path(filename):
    """Resolve a writable config path for both source and packaged runs."""
    candidate_dirs = []
    local_appdata = os.environ.get("LOCALAPPDATA")
    if local_appdata:
        candidate_dirs.append(os.path.join(local_appdata, "AcademicFigureCropper"))

    home_dir = os.path.expanduser("~")
    if home_dir:
        candidate_dirs.append(os.path.join(home_dir, ".academic_figure_cropper"))

    candidate_dirs.append(os.path.abspath("."))

    for directory in candidate_dirs:
        try:
            os.makedirs(directory, exist_ok=True)
            probe_path = os.path.join(directory, ".write_test")
            with open(probe_path, "w", encoding="utf-8") as probe_file:
                probe_file.write("ok")
            os.remove(probe_path)
            return os.path.join(directory, filename)
        except OSError:
            continue

    return os.path.join(os.path.abspath("."), filename)

def enable_high_dpi():
    """Try to make the process DPI-aware on Windows to avoid bitmap-scaled blurry UI."""
    if not sys.platform.startswith("win"):
        return

    try:
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
        return
    except Exception:
        pass

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        return
    except Exception:
        pass

    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

class PDFCropperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Academic Figure Cropper")
        self.root.geometry("420x370")
        self.root.minsize(390, 320)
        self.root.resizable(False, False)
        
        # 设置窗口图标（如果有的话）
        try:
            icon_path = resource_path("icon.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"加载图标失败: {e}")
            pass
        
        # 设置主题色
        self.primary_color = "#2563eb"
        self.primary_soft_color = "#dbeafe"
        self.primary_hover_color = "#1d4ed8"
        self.success_color = "#16a34a"
        self.success_soft_color = "#dcfce7"
        self.warning_color = "#ea580c"
        self.warning_soft_color = "#ffedd5"
        self.bg_color = "#eff3f8"
        self.card_bg_color = "#ffffff"
        self.muted_bg_color = "#f8fafc"
        self.text_color = "#0f172a"
        self.secondary_text = "#64748b"
        self.button_text_color = "#ffffff"
        self.border_color = "#d8e1eb"
        self.drop_border_color = "#c7d7ea"
        self.disabled_text_color = "#94a3b8"
        self.font_family = "Microsoft YaHei UI" if sys.platform.startswith("win") else "Segoe UI"
        self.title_font = (self.font_family, 11, "bold")
        self.body_font = (self.font_family, 10)
        self.small_font = (self.font_family, 9)
        self.drop_title_font = (self.font_family, 15, "bold")
        self.badge_font = (self.font_family, 9, "bold")
        
        # 支持的图片格式
        self.supported_img_formats = ('.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.gif')
        
        # 配置样式
        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass
        self.style.configure(
            "Slim.Horizontal.TProgressbar",
            background=self.primary_color,
            troughcolor=self.muted_bg_color,
            bordercolor=self.muted_bg_color,
            lightcolor=self.primary_color,
            darkcolor=self.primary_color,
            thickness=8,
        )
        
        # 设置背景色
        self.root.configure(bg=self.bg_color)
        
        # 创建配置文件管理
        config_name = "pdf_cropper_config.ini"
        self.config_file = get_config_path(config_name)
        self.config = configparser.ConfigParser()
        self.load_config()
        self.save_debug_images = self.config.getboolean('Settings', 'save_debug_images')
        self.is_processing = False
        self.advanced_visible = False
        self.last_output_dir = ""
        self._layout_update_job = None
        self._pending_canvas_width = None
        self._last_canvas_width = None
        self.root.attributes("-topmost", self.config.getboolean('Settings', 'always_on_top'))
        
        # 进度变量
        self.progress_var = tk.DoubleVar()
        self.progress_var.set(0)
        self.ui_queue = queue.Queue()
        
        # 创建UI元素
        self.create_widgets()
        
        # 处理的文件列表
        self.processing_files = []
        self.root.after(50, self.process_ui_queue)
        
    def load_config(self):
        """加载配置文件"""
        if os.path.exists(self.config_file):
            self.config.read(self.config_file)
        
        if 'Settings' not in self.config:
            self.config['Settings'] = {}
            
        if 'overwrite_original' not in self.config['Settings']:
            self.config['Settings']['overwrite_original'] = 'True'
            
        if 'output_dir' not in self.config['Settings']:
            self.config['Settings']['output_dir'] = ''
            
        if 'left_margin' not in self.config['Settings']:
            self.config['Settings']['left_margin'] = '0'
            
        if 'right_margin' not in self.config['Settings']:
            self.config['Settings']['right_margin'] = '0'
            
        if 'top_margin' not in self.config['Settings']:
            self.config['Settings']['top_margin'] = '0'
            
        if 'bottom_margin' not in self.config['Settings']:
            self.config['Settings']['bottom_margin'] = '0'

        if 'always_on_top' not in self.config['Settings']:
            self.config['Settings']['always_on_top'] = 'True'

        if 'save_debug_images' not in self.config['Settings']:
            self.config['Settings']['save_debug_images'] = 'False'
    
    def save_config(self):
        """保存配置到文件"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                self.config.write(f)
        except OSError as exc:
            print(f"保存配置失败: {exc}")
    
    def create_widgets(self):
        """创建UI元素"""
        self.container = tk.Frame(self.root, bg=self.bg_color)
        self.container.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        self.header_frame = tk.Frame(self.container, bg=self.bg_color)
        self.header_frame.pack(fill=tk.X)

        self.title_frame = tk.Frame(self.header_frame, bg=self.bg_color)
        self.title_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.title_label = tk.Label(
            self.title_frame,
            text="Academic Figure Cropper",
            font=self.title_font,
            fg=self.text_color,
            bg=self.bg_color,
        )
        self.title_label.pack(anchor=tk.W)
        self.subtitle_label = tk.Label(
            self.title_frame,
            text="拖入文件即可自动裁白边",
            font=self.small_font,
            fg=self.secondary_text,
            bg=self.bg_color,
        )
        self.subtitle_label.pack(anchor=tk.W, pady=(2, 0))

        header_actions = tk.Frame(self.header_frame, bg=self.bg_color)
        header_actions.pack(side=tk.RIGHT)

        self.topmost_var = tk.BooleanVar(value=self.config.getboolean('Settings', 'always_on_top'))
        self.topmost_button = self.create_flat_button(header_actions, "", self.toggle_topmost, compact=True)
        self.topmost_button.pack(side=tk.LEFT)
        self.bind_window_drag([self.header_frame, self.title_frame, self.title_label, self.subtitle_label])

        self.scroll_host = tk.Frame(self.container, bg=self.bg_color)
        self.scroll_host.pack(fill=tk.BOTH, expand=True, pady=(12, 0))

        self.scroll_canvas = tk.Canvas(
            self.scroll_host,
            bg=self.bg_color,
            highlightthickness=0,
            bd=0,
            relief=tk.FLAT,
        )
        self.scroll_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.content_frame = tk.Frame(self.scroll_canvas, bg=self.bg_color)
        self.canvas_window = self.scroll_canvas.create_window((0, 0), window=self.content_frame, anchor="nw")

        self.content_frame.bind("<Configure>", self.on_frame_configure)
        self.scroll_canvas.bind("<Configure>", self.on_canvas_configure)
        self.root.bind_all("<MouseWheel>", self.on_mousewheel)
        self.root.bind_all("<Button-4>", self.on_mousewheel)
        self.root.bind_all("<Button-5>", self.on_mousewheel)

        self.drop_card = tk.Frame(
            self.content_frame,
            bg=self.card_bg_color,
            highlightthickness=1,
            highlightbackground=self.drop_border_color,
            highlightcolor=self.drop_border_color,
        )
        self.drop_card.pack(fill=tk.BOTH, expand=True, pady=(12, 10))

        self.drop_body = tk.Frame(self.drop_card, bg=self.card_bg_color, padx=16, pady=14)
        self.drop_body.pack(fill=tk.BOTH, expand=True)

        drop_center_frame = tk.Frame(self.drop_body, bg=self.card_bg_color)
        drop_center_frame.pack(fill=tk.BOTH, expand=True)

        self.drop_badge = tk.Label(
            drop_center_frame,
            text="DROP",
            font=self.badge_font,
            fg=self.primary_color,
            bg=self.primary_soft_color,
            padx=10,
            pady=4,
        )
        self.drop_badge.pack(anchor=tk.CENTER, pady=(4, 10))

        self.drop_label = tk.Label(
            drop_center_frame,
            text="拖入 PDF 或图片",
            font=self.drop_title_font,
            fg=self.text_color,
            bg=self.card_bg_color,
        )
        self.drop_label.pack()

        self.drop_hint = tk.Label(
            drop_center_frame,
            text="自动裁白边并保存",
            font=self.body_font,
            fg=self.secondary_text,
            bg=self.card_bg_color,
        )
        self.drop_hint.pack(pady=(6, 12))

        self.pick_button = self.create_flat_button(drop_center_frame, "选择文件", self.select_files, primary=True)
        self.pick_button.pack()

        status_frame = tk.Frame(self.drop_body, bg=self.card_bg_color)
        status_frame.pack(fill=tk.X, pady=(14, 0))

        self.status_var = tk.StringVar(value="等待文件")
        self.status_label = tk.Label(
            status_frame,
            textvariable=self.status_var,
            font=self.small_font,
            fg=self.secondary_text,
            bg=self.card_bg_color,
        )
        self.status_label.pack(anchor=tk.W)

        self.progress = ttk.Progressbar(
            status_frame,
            variable=self.progress_var,
            mode='determinate',
            style="Slim.Horizontal.TProgressbar",
        )
        self.progress.pack(fill=tk.X, pady=(8, 0))

        self.bind_drop_targets([
            self.drop_card,
            self.drop_body,
            drop_center_frame,
            self.drop_badge,
            self.drop_label,
            self.drop_hint,
        ])

        self.toolbar_card = tk.Frame(
            self.content_frame,
            bg=self.card_bg_color,
            highlightthickness=1,
            highlightbackground=self.border_color,
            highlightcolor=self.border_color,
        )
        self.toolbar_card.pack(fill=tk.X)

        toolbar_body = tk.Frame(self.toolbar_card, bg=self.card_bg_color, padx=10, pady=10)
        toolbar_body.pack(fill=tk.X)

        mode_frame = tk.Frame(toolbar_body, bg=self.card_bg_color)
        mode_frame.pack(fill=tk.X)

        self.overwrite_var = tk.BooleanVar(value=self.config.getboolean('Settings', 'overwrite_original'))
        self.overwrite_button = self.create_flat_button(
            mode_frame,
            "覆盖原文件",
            lambda: self.set_output_mode(True),
            compact=True,
        )
        self.overwrite_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

        self.output_dir_button = self.create_flat_button(
            mode_frame,
            "输出到目录",
            lambda: self.set_output_mode(False),
            compact=True,
        )
        self.output_dir_button.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.output_path_var = tk.StringVar(value=self.config.get('Settings', 'output_dir'))
        self.output_path_label = tk.Label(
            toolbar_body,
            text="输出目录",
            font=self.small_font,
            fg=self.secondary_text,
            bg=self.card_bg_color,
        )
        self.output_path_label.pack(anchor=tk.W, pady=(10, 6))

        self.output_path_row = tk.Frame(toolbar_body, bg=self.card_bg_color)
        self.output_path_row.pack(fill=tk.X)

        self.output_entry = tk.Entry(
            self.output_path_row,
            textvariable=self.output_path_var,
            font=self.small_font,
            relief=tk.FLAT,
            bd=0,
            bg=self.muted_bg_color,
            fg=self.text_color,
            insertbackground=self.text_color,
            highlightthickness=1,
            highlightbackground=self.border_color,
            highlightcolor=self.primary_color,
            disabledbackground=self.muted_bg_color,
            disabledforeground=self.disabled_text_color,
        )
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8)
        self.output_entry.bind("<FocusOut>", self.persist_output_path)
        self.output_entry.bind("<Return>", self.persist_output_path)

        self.output_browse_button = self.create_flat_button(
            self.output_path_row,
            "浏览",
            self.select_output_dir,
            compact=True,
        )
        self.output_browse_button.pack(side=tk.LEFT, padx=(8, 6))

        self.output_open_button = self.create_flat_button(
            self.output_path_row,
            "打开",
            self.open_output_dir,
            compact=True,
        )
        self.output_open_button.pack(side=tk.LEFT)

        self.output_mode_hint_var = tk.StringVar(value="")
        self.output_mode_hint_label = tk.Label(
            toolbar_body,
            textvariable=self.output_mode_hint_var,
            font=self.small_font,
            fg=self.secondary_text,
            bg=self.card_bg_color,
        )
        self.output_mode_hint_label.pack(anchor=tk.W, pady=(6, 0))

        margin_row = tk.Frame(toolbar_body, bg=self.card_bg_color)
        margin_row.pack(fill=tk.X, pady=(10, 0))

        tk.Label(
            margin_row,
            text="留白",
            font=self.small_font,
            fg=self.secondary_text,
            bg=self.card_bg_color,
        ).pack(side=tk.LEFT)

        self.uniform_margin_var = tk.StringVar()
        self.uniform_margin_spin = tk.Spinbox(
            margin_row,
            from_=0,
            to=50,
            width=5,
            textvariable=self.uniform_margin_var,
            command=self.apply_uniform_margin,
            font=self.small_font,
            justify="center",
            relief=tk.FLAT,
            bd=0,
            bg=self.muted_bg_color,
            fg=self.text_color,
            buttonbackground=self.muted_bg_color,
            insertbackground=self.text_color,
            highlightthickness=1,
            highlightbackground=self.border_color,
            highlightcolor=self.primary_color,
        )
        self.uniform_margin_spin.pack(side=tk.LEFT, padx=(8, 6), ipady=5)
        self.uniform_margin_spin.bind("<FocusOut>", self.apply_uniform_margin)
        self.uniform_margin_spin.bind("<Return>", self.apply_uniform_margin)

        tk.Label(
            margin_row,
            text="px",
            font=self.small_font,
            fg=self.secondary_text,
            bg=self.card_bg_color,
        ).pack(side=tk.LEFT)

        self.margin_details_button = self.create_flat_button(
            margin_row,
            "",
            self.toggle_advanced_options,
            compact=True,
        )
        self.margin_details_button.pack(side=tk.RIGHT)

        self.margin_summary_var = tk.StringVar(value="")
        self.margin_summary_label = tk.Label(
            toolbar_body,
            textvariable=self.margin_summary_var,
            font=self.small_font,
            fg=self.secondary_text,
            bg=self.card_bg_color,
        )
        self.margin_summary_label.pack(anchor=tk.W, pady=(6, 0))

        self.advanced_panel = tk.Frame(
            self.content_frame,
            bg=self.card_bg_color,
            highlightthickness=1,
            highlightbackground=self.border_color,
            highlightcolor=self.border_color,
        )

        advanced_body = tk.Frame(self.advanced_panel, bg=self.card_bg_color, padx=12, pady=12)
        advanced_body.pack(fill=tk.X)

        tk.Label(
            advanced_body,
            text="额外留白 (px)",
            font=self.body_font,
            fg=self.text_color,
            bg=self.card_bg_color,
        ).pack(anchor=tk.W)
        tk.Label(
            advanced_body,
            text="裁剪后额外保留的空白，0 表示贴边裁剪。",
            font=self.small_font,
            fg=self.secondary_text,
            bg=self.card_bg_color,
        ).pack(anchor=tk.W, pady=(2, 10))

        margins_grid = tk.Frame(advanced_body, bg=self.card_bg_color)
        margins_grid.pack(fill=tk.X)
        margins_grid.grid_columnconfigure(0, weight=1)
        margins_grid.grid_columnconfigure(1, weight=1)

        self.left_margin_var = tk.IntVar(value=int(self.config.get('Settings', 'left_margin')))
        self.right_margin_var = tk.IntVar(value=int(self.config.get('Settings', 'right_margin')))
        self.top_margin_var = tk.IntVar(value=int(self.config.get('Settings', 'top_margin')))
        self.bottom_margin_var = tk.IntVar(value=int(self.config.get('Settings', 'bottom_margin')))

        self.left_margin_spin = self.create_margin_field(margins_grid, "左", self.left_margin_var, 0, 0)
        self.right_margin_spin = self.create_margin_field(margins_grid, "右", self.right_margin_var, 0, 1)
        self.top_margin_spin = self.create_margin_field(margins_grid, "上", self.top_margin_var, 1, 0)
        self.bottom_margin_spin = self.create_margin_field(margins_grid, "下", self.bottom_margin_var, 1, 1)

        self.update_topmost_button()
        self.update_advanced_button()
        self.toggle_output_path()
        self.set_drop_area_state("idle")
        self.root.after_idle(self.delayed_layout_update)

    def create_flat_button(self, parent, text, command, primary=False, compact=False):
        font = self.small_font if compact else self.body_font
        button = tk.Button(
            parent,
            text=text,
            command=command,
            font=font,
            relief=tk.FLAT,
            bd=0,
            cursor="hand2",
            padx=12,
            pady=6 if compact else 8,
            disabledforeground=self.disabled_text_color,
        )
        if primary:
            button.configure(
                bg=self.primary_color,
                fg=self.button_text_color,
                activebackground=self.primary_hover_color,
                activeforeground=self.button_text_color,
                highlightthickness=1,
                highlightbackground=self.primary_color,
                highlightcolor=self.primary_hover_color,
            )
        else:
            button.configure(
                bg=self.card_bg_color,
                fg=self.text_color,
                activebackground=self.muted_bg_color,
                activeforeground=self.text_color,
                highlightthickness=1,
                highlightbackground=self.border_color,
                highlightcolor=self.border_color,
            )
        return button

    def create_margin_field(self, parent, label_text, variable, row, column):
        field_frame = tk.Frame(parent, bg=self.card_bg_color)
        field_frame.grid(row=row, column=column, sticky="ew", padx=(0, 8) if column == 0 else (8, 0), pady=(0, 8) if row == 0 else (8, 0))

        tk.Label(
            field_frame,
            text=label_text,
            font=self.small_font,
            fg=self.secondary_text,
            bg=self.card_bg_color,
        ).pack(anchor=tk.W, pady=(0, 4))

        spinbox = tk.Spinbox(
            field_frame,
            from_=0,
            to=50,
            width=6,
            textvariable=variable,
            command=self.save_margins,
            font=self.small_font,
            justify="center",
            relief=tk.FLAT,
            bd=0,
            bg=self.muted_bg_color,
            fg=self.text_color,
            buttonbackground=self.muted_bg_color,
            insertbackground=self.text_color,
            highlightthickness=1,
            highlightbackground=self.border_color,
            highlightcolor=self.primary_color,
        )
        spinbox.pack(fill=tk.X, ipady=6)
        spinbox.bind("<FocusOut>", self.save_margins)
        spinbox.bind("<Return>", self.save_margins)
        return spinbox

    def bind_drop_targets(self, widgets):
        for widget in widgets:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind('<<Drop>>', self.drop)
            widget.dnd_bind('<<DropEnter>>', self.on_drop_enter)
            widget.dnd_bind('<<DropLeave>>', self.on_drop_leave)

    def bind_window_drag(self, widgets):
        for widget in widgets:
            widget.bind("<ButtonPress-1>", self.start_window_drag)
            widget.bind("<B1-Motion>", self.on_window_drag)
            widget.bind("<ButtonRelease-1>", self.stop_window_drag)

    def start_window_drag(self, event):
        self._drag_offset_x = event.x_root - self.root.winfo_x()
        self._drag_offset_y = event.y_root - self.root.winfo_y()

    def on_window_drag(self, event):
        if not hasattr(self, "_drag_offset_x"):
            return

        x = event.x_root - self._drag_offset_x
        y = event.y_root - self._drag_offset_y
        self.root.geometry(f"+{x}+{y}")

    def stop_window_drag(self, event):
        self._drag_offset_x = None
        self._drag_offset_y = None

    def set_output_mode(self, overwrite_original):
        self.overwrite_var.set(overwrite_original)
        self.toggle_output_path()

    def update_chip_button(self, button, selected):
        if selected:
            button.config(
                bg=self.primary_color,
                fg=self.button_text_color,
                activebackground=self.primary_hover_color,
                activeforeground=self.button_text_color,
                highlightbackground=self.primary_color,
                highlightcolor=self.primary_hover_color,
            )
        else:
            button.config(
                bg=self.card_bg_color,
                fg=self.text_color,
                activebackground=self.muted_bg_color,
                activeforeground=self.text_color,
                highlightbackground=self.border_color,
                highlightcolor=self.border_color,
            )

    def update_topmost_button(self):
        enabled = self.topmost_var.get()
        self.update_chip_button(self.topmost_button, enabled)
        self.topmost_button.config(text=f"置顶 {'开' if enabled else '关'}")

    def get_margin_values(self):
        return [
            self.left_margin_var.get(),
            self.right_margin_var.get(),
            self.top_margin_var.get(),
            self.bottom_margin_var.get(),
        ]

    def get_uniform_margin_display_value(self):
        margins = self.get_margin_values()
        return str(margins[0]) if len(set(margins)) == 1 else ""

    def toggle_topmost(self):
        enabled = not self.topmost_var.get()
        self.topmost_var.set(enabled)
        self.root.attributes("-topmost", enabled)
        self.config['Settings']['always_on_top'] = str(enabled)
        self.save_config()
        self.update_topmost_button()

    def update_advanced_button(self):
        margins = self.get_margin_values() if hasattr(self, 'left_margin_var') else [0, 0, 0, 0]
        if all(margin == 0 for margin in margins):
            self.margin_summary_var.set("0 表示贴边裁剪，不额外保留空白。")
        elif len(set(margins)) == 1:
            self.margin_summary_var.set(f"当前四边统一留白 {margins[0]}px。")
        else:
            self.margin_summary_var.set(
                f"当前分别设置: 左{margins[0]} 右{margins[1]} 上{margins[2]} 下{margins[3]} px"
            )

        if hasattr(self, 'uniform_margin_var'):
            self.uniform_margin_var.set(self.get_uniform_margin_display_value())

        button_text = f"分别设置 {'▲' if self.advanced_visible else '▼'}"
        self.update_chip_button(self.margin_details_button, self.advanced_visible)
        self.margin_details_button.config(text=button_text)

    def toggle_advanced_options(self):
        self.advanced_visible = not self.advanced_visible
        if self.advanced_visible:
            self.advanced_panel.pack(fill=tk.X, pady=(10, 0))
        else:
            self.advanced_panel.pack_forget()
        self.update_advanced_button()
        self.root.after_idle(self.delayed_layout_update)
        if self.advanced_visible:
            self.root.after_idle(lambda: self.scroll_canvas.yview_moveto(1.0))

    def apply_uniform_margin(self, event=None):
        raw_value = self.uniform_margin_var.get().strip()
        if raw_value == "":
            return

        try:
            margin_value = int(raw_value)
        except ValueError:
            self.uniform_margin_var.set(self.get_uniform_margin_display_value())
            return

        margin_value = max(0, min(50, margin_value))
        self.left_margin_var.set(margin_value)
        self.right_margin_var.set(margin_value)
        self.top_margin_var.set(margin_value)
        self.bottom_margin_var.set(margin_value)
        self.save_margins()

    def toggle_output_path(self):
        """根据覆盖选项切换输出路径控件"""
        overwrite_original = self.overwrite_var.get()
        self.config['Settings']['overwrite_original'] = str(overwrite_original)
        self.save_config()

        self.update_chip_button(self.overwrite_button, overwrite_original)
        self.update_chip_button(self.output_dir_button, not overwrite_original)
        self.output_entry.config(state=tk.DISABLED if overwrite_original else tk.NORMAL)
        self.output_browse_button.config(state=tk.DISABLED if overwrite_original else tk.NORMAL)

        self.update_output_path_buttons()
        self.root.after_idle(self.delayed_layout_update)

    def update_output_path_buttons(self):
        output_dir = self.output_path_var.get().strip()
        has_directory = bool(output_dir) and os.path.isdir(output_dir)
        overwrite_original = self.overwrite_var.get()
        self.output_open_button.config(state=tk.NORMAL if has_directory and not overwrite_original else tk.DISABLED)

        if overwrite_original:
            self.output_mode_hint_var.set("当前会直接覆盖原文件，无需设置输出目录。")
        elif output_dir:
            self.output_mode_hint_var.set("当前会保存到上面的输出目录。")
        else:
            self.output_mode_hint_var.set("点击“浏览”选择一个输出目录。")

    def persist_output_path(self, event=None):
        output_dir = self.output_path_var.get().strip()
        self.output_path_var.set(output_dir)
        self.config['Settings']['output_dir'] = output_dir
        self.save_config()
        self.update_output_path_buttons()

    def save_margins(self, event=None):
        """保存边距设置"""
        self.config['Settings']['left_margin'] = str(self.left_margin_var.get())
        self.config['Settings']['right_margin'] = str(self.right_margin_var.get())
        self.config['Settings']['top_margin'] = str(self.top_margin_var.get())
        self.config['Settings']['bottom_margin'] = str(self.bottom_margin_var.get())
        self.save_config()
        self.update_advanced_button()

    def select_files(self):
        file_types = [
            ("支持的文件", "*.pdf *.jpg *.jpeg *.png *.bmp *.tiff *.tif *.gif"),
            ("PDF 文件", "*.pdf"),
            ("图片文件", "*.jpg *.jpeg *.png *.bmp *.tiff *.tif *.gif"),
            ("所有文件", "*.*"),
        ]
        files = filedialog.askopenfilenames(title="选择要处理的文件", filetypes=file_types)
        if files:
            self.process_dropped_files(list(files))

    def select_output_dir(self):
        """选择输出目录"""
        initial_dir = self.output_path_var.get().strip() or os.path.expanduser("~")
        output_dir = filedialog.askdirectory(initialdir=initial_dir)
        if output_dir:
            self.output_path_var.set(output_dir)
            self.persist_output_path()

    def open_output_dir(self):
        output_dir = self.output_path_var.get().strip()
        if not output_dir:
            messagebox.showwarning("提示", "请先设置输出文件夹")
            return

        if not os.path.isdir(output_dir):
            messagebox.showwarning("提示", "输出文件夹不存在")
            return

        try:
            if hasattr(os, "startfile"):
                os.startfile(output_dir)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", output_dir])
            else:
                subprocess.Popen(["xdg-open", output_dir])
        except Exception as exc:
            messagebox.showerror("错误", f"无法打开输出文件夹:\n{exc}")

    def set_drop_area_state(self, state):
        palette = {
            "idle": (self.drop_border_color, self.primary_soft_color, self.primary_color, "DROP"),
            "drag": (self.primary_color, self.primary_soft_color, self.primary_color, "READY"),
            "processing": (self.primary_color, self.primary_soft_color, self.primary_color, "WORK"),
            "success": (self.success_color, self.success_soft_color, self.success_color, "DONE"),
            "warning": (self.warning_color, self.warning_soft_color, self.warning_color, "WARN"),
        }
        border_color, badge_bg, badge_fg, badge_text = palette[state]
        self.drop_card.config(highlightbackground=border_color, highlightcolor=border_color)
        self.drop_badge.config(bg=badge_bg, fg=badge_fg, text=badge_text)

    def on_drop_enter(self, event):
        if not self.is_processing:
            self.set_drop_area_state("drag")
        return event.action

    def on_drop_leave(self, event):
        if not self.is_processing:
            self.set_drop_area_state("idle")

    def get_processing_settings(self):
        """Capture UI-controlled settings before starting a worker thread."""
        output_dir = self.output_path_var.get().strip()
        self.config['Settings']['output_dir'] = output_dir
        self.save_config()
        return {
            'overwrite_original': self.overwrite_var.get(),
            'output_dir': output_dir,
            'margins': {
                'left': self.left_margin_var.get(),
                'right': self.right_margin_var.get(),
                'top': self.top_margin_var.get(),
                'bottom': self.bottom_margin_var.get(),
            },
            'save_debug_images': self.save_debug_images,
        }

    def enqueue_ui_call(self, callback, *args, **kwargs):
        self.ui_queue.put((callback, args, kwargs))

    def process_ui_queue(self):
        try:
            while True:
                callback, args, kwargs = self.ui_queue.get_nowait()
                callback(*args, **kwargs)
        except queue.Empty:
            pass
        finally:
            self.root.after(50, self.process_ui_queue)

    def build_output_path(self, file_path, settings, reserved_paths):
        if settings['overwrite_original']:
            return file_path

        output_dir = settings['output_dir']
        output_name = os.path.basename(file_path)
        base_name, ext = os.path.splitext(output_name)
        candidate = os.path.join(output_dir, f"{base_name}_cropped{ext}")
        candidate_key = os.path.normcase(os.path.abspath(candidate))
        suffix = 2

        while candidate_key in reserved_paths or os.path.exists(candidate):
            candidate = os.path.join(output_dir, f"{base_name}_cropped_{suffix}{ext}")
            candidate_key = os.path.normcase(os.path.abspath(candidate))
            suffix += 1

        reserved_paths.add(candidate_key)
        return candidate

    def finish_processing(self, total_success, total_failed, failed_messages):
        self.is_processing = False

        if total_failed > 0:
            self.status_var.set(f"处理完成: {total_success} 成功, {total_failed} 失败")
            self.status_label.config(fg=self.warning_color)
            self.set_drop_area_state("warning")
            summary = "\n".join(failed_messages[:6])
            if len(failed_messages) > 6:
                summary += f"\n... 另有 {len(failed_messages) - 6} 个文件失败"
            messagebox.showwarning("处理完成", f"成功: {total_success} 个文件\n失败: {total_failed} 个文件\n\n{summary}")
        else:
            self.status_var.set(f"已完成 {total_success} 个文件")
            self.status_label.config(fg=self.success_color)
            self.set_drop_area_state("success")
    
    def drop(self, event):
        """处理文件拖放事件"""
        if not self.is_processing:
            self.set_drop_area_state("idle")
        files = self.parse_drop_data(event.data)
        self.process_dropped_files(files)
    
    def parse_drop_data(self, data):
        """解析拖放的文件路径数据"""
        files = []
        for item in self.root.tk.splitlist(data):
            # 处理可能的引号和花括号（Windows路径特性）
            item = item.strip('{}')
            # 检查文件扩展名
            _, ext = os.path.splitext(item.lower())
            if ext == '.pdf' or ext in self.supported_img_formats:
                files.append(item)
        return files
    
    def process_dropped_files(self, files):
        """处理拖放的文件"""
        if self.is_processing:
            messagebox.showinfo("请稍候", "当前还有文件在处理中")
            return

        if not files:
            self.status_var.set("没有检测到支持的文件")
            self.status_label.config(fg=self.warning_color)
            self.set_drop_area_state("warning")
            messagebox.showinfo("提示", "没有检测到支持的文件格式")
            return

        settings = self.get_processing_settings()

        # 检查输出路径
        if not settings['overwrite_original'] and not settings['output_dir']:
            self.status_var.set("请先设置输出文件夹")
            self.status_label.config(fg=self.warning_color)
            self.set_drop_area_state("warning")
            messagebox.showwarning("警告", "请先选择输出文件夹")
            return

        if not settings['overwrite_original']:
            try:
                os.makedirs(settings['output_dir'], exist_ok=True)
            except OSError as exc:
                self.status_var.set("输出文件夹不可用")
                self.status_label.config(fg=self.warning_color)
                self.set_drop_area_state("warning")
                messagebox.showerror("错误", f"无法创建输出文件夹:\n{exc}")
                return

        self.update_output_path_buttons()
        self.is_processing = True
        self.last_output_dir = settings['output_dir'] if not settings['overwrite_original'] else ""
        self.progress_var.set(0)
        self.progress.config(maximum=len(files))

        # 更新UI反馈
        self.status_label.config(fg=self.secondary_text)
        self.set_drop_area_state("processing")
        self.status_var.set(f"开始处理 {len(files)} 个文件...")

        # 开始处理线程
        threading.Thread(target=self.process_files_thread, args=(files, settings), daemon=True).start()

    def process_files_thread(self, files, settings):
        """在单独的线程中处理文件"""
        total_success = 0
        total_failed = 0
        reserved_output_paths = set()
        failed_messages = []

        for i, file_path in enumerate(files):
            try:
                filename = os.path.basename(file_path)
                self.enqueue_ui_call(self.status_var.set, f"正在处理 {i + 1}/{len(files)} · {filename}")

                # 确定输出路径
                output_path = self.build_output_path(file_path, settings, reserved_output_paths)

                # 根据文件类型选择处理方法
                _, ext = os.path.splitext(file_path.lower())
                if ext == '.pdf':
                    self.crop_pdf(file_path, output_path, settings)
                elif ext in self.supported_img_formats:
                    self.crop_image(file_path, output_path, settings['margins'])

                # 更新进度
                self.enqueue_ui_call(self.progress_var.set, i + 1)
                total_success += 1

            except Exception as e:
                total_failed += 1
                failed_messages.append(f"{os.path.basename(file_path)}: {str(e)}")

        self.enqueue_ui_call(self.finish_processing, total_success, total_failed, failed_messages)

    def crop_image(self, input_path, output_path, margins):
        """剪裁图片白边"""
        # 打开图片
        img = Image.open(input_path)
        
        # 确保图片是RGB模式，以便于处理
        if img.mode != 'RGB':
            img = img.convert('RGB')
        
        # 将图片转换为numpy数组
        np_img = np.array(img)
        
        # 计算图片亮度
        brightness = np.mean(np_img, axis=2)
        
        # 使用阈值确定非白色像素
        threshold = 225
        mask = brightness < threshold
        
        # 找到内容区域边界
        if np.any(mask):  # 如果有任何内容
            rows = np.where(np.any(mask, axis=1))[0]
            cols = np.where(np.any(mask, axis=0))[0]
            
            if len(rows) > 0 and len(cols) > 0:
                min_y, max_y = np.min(rows), np.max(rows)
                min_x, max_x = np.min(cols), np.max(cols)
                
                # 噪点过滤
                min_content_size = 10
                if (max_x - min_x) > min_content_size and (max_y - min_y) > min_content_size:
                    # 获取边距设置
                    left_margin = margins['left']
                    top_margin = margins['top']
                    right_margin = margins['right']
                    bottom_margin = margins['bottom']
                    
                    # 计算裁剪区域（添加边距）
                    x1 = max(min_x - left_margin, 0)
                    y1 = max(min_y - top_margin, 0)
                    x2 = min(max_x + right_margin, np_img.shape[1])
                    y2 = min(max_y + bottom_margin, np_img.shape[0])
                    
                    # 内容区域有效性验证
                    width, height = np_img.shape[1], np_img.shape[0]
                    
                    # 防止裁剪过多 - 如果内容区域太小，可能是错误检测
                    if (x2 - x1) < width * 0.1 or (y2 - y1) < height * 0.1:
                        x1, y1, x2, y2 = 0, 0, width, height
                    
                    # 防止裁剪过少 - 如果内容区域几乎和页面一样大，微调一下裁剪区域
                    if (x2 - x1) > width * 0.98 or (y2 - y1) > height * 0.98:
                        margin_x = width * 0.02
                        margin_y = height * 0.02
                        x1, y1 = margin_x, margin_y
                        x2, y2 = width - margin_x, height - margin_y
                    
                    # 裁剪图片
                    cropped_img = img.crop((x1, y1, x2, y2))
                    
                    # 保存裁剪后的图片
                    if input_path == output_path:
                        # 如果覆盖原文件，先保存为临时文件再替换
                        temp_path = output_path + ".temp"
                        # 获取原文件的扩展名
                        _, ext = os.path.splitext(input_path)
                        # 确保临时文件保留原始扩展名
                        cropped_img.save(temp_path, format=self.get_image_format(ext))
                        cropped_img.close()
                        img.close()
                        os.replace(temp_path, output_path)
                    else:
                        # 直接保存到新位置
                        _, ext = os.path.splitext(output_path)
                        cropped_img.save(output_path, format=self.get_image_format(ext))
                        cropped_img.close()
                        img.close()
                    return
        
        # 如果没有检测到内容或检测失败，保存原图
        if input_path != output_path:
            _, ext = os.path.splitext(output_path)
            img.save(output_path, format=self.get_image_format(ext))
        img.close()
    
    def get_image_format(self, ext):
        """根据文件扩展名获取图片格式"""
        ext = ext.lower().strip('.')
        # 处理特殊情况
        if ext == 'jpg':
            return 'JPEG'
        elif ext == 'tif':
            return 'TIFF'
        elif ext in ('jpeg', 'png', 'bmp', 'tiff', 'gif'):
            return ext.upper()
        # 默认返回PNG格式
        return 'PNG'
    
    def crop_pdf(self, input_path, output_path, settings):
        """剪裁PDF文件白边"""
        # 打开PDF文件
        doc = fitz.open(input_path)
        
        # 创建新文档用于保存
        new_doc = fitz.open()
        
        # 获取边距设置
        margins = settings['margins']
        left_margin = margins['left']
        top_margin = margins['top']
        right_margin = margins['right']
        bottom_margin = margins['bottom']

        # 创建调试输出目录
        debug_dir = None
        if settings.get('save_debug_images'):
            debug_dir = os.path.join(os.path.dirname(output_path), "debug_output")
            os.makedirs(debug_dir, exist_ok=True)
        
        # 处理每一页
        for page_num in range(len(doc)):
            try:
                page = doc.load_page(page_num)
                
                # 获取页面的边界框
                rect = page.rect
                
                # 使用pixmap分析页面内容
                try:
                    # 提高分辨率以获取更精确的边界
                    pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False)  # 增加到10倍采样
                    img_data = pix.samples
                    width, height = pix.width, pix.height
                    
                    # 使用numpy处理图像数据
                    channels = pix.n
                    # 将原始数据转换为numpy数组
                    np_img = np.frombuffer(img_data, dtype=np.uint8).reshape(height, width, channels)
                    
                    # 保存原始图像用于调试
                    if debug_dir:
                        debug_img_path = os.path.join(debug_dir, f"page_{page_num+1}_original.png")
                        Image.fromarray(np_img).save(debug_img_path)
                    
                    # 计算亮度 - 对于RGB图像取平均值，对于灰度图像直接使用值
                    if channels >= 3:
                        # 对于RGB或RGBA图像
                        brightness = np.mean(np_img[:, :, :3], axis=2)
                    else:
                        # 对于灰度图像
                        brightness = np_img[:, :, 0]
                    
                    # 保存亮度图用于调试
                    if debug_dir:
                        debug_brightness_path = os.path.join(debug_dir, f"page_{page_num+1}_brightness.png")
                        Image.fromarray(brightness.astype(np.uint8)).save(debug_brightness_path)
                    
                    # 根据阈值创建掩码
                    threshold = 245  # 阈值
                    mask = brightness < threshold
                    
                    # 保存掩码图用于调试
                    if debug_dir:
                        debug_mask_path = os.path.join(debug_dir, f"page_{page_num+1}_mask.png")
                        Image.fromarray((mask * 255).astype(np.uint8)).save(debug_mask_path)
                    
                    # 根据掩码找到内容区域边界
                    if np.any(mask):  # 如果有任何内容
                        # 从四个方向向中间扫描
                        # 从左向右扫描
                        left_bound = 0
                        for x in range(width):
                            if np.any(mask[:, x]):
                                left_bound = x
                                break
                        
                        # 从右向左扫描
                        right_bound = width - 1
                        for x in range(width-1, -1, -1):
                            if np.any(mask[:, x]):
                                right_bound = x
                                break
                        
                        # 从上向下扫描
                        top_bound = 0
                        for y in range(height):
                            if np.any(mask[y, :]):
                                top_bound = y
                                break
                        
                        # 从下向上扫描
                        bottom_bound = height - 1
                        for y in range(height-1, -1, -1):
                            if np.any(mask[y, :]):
                                bottom_bound = y
                                break
                        
                        # 将像素坐标转换回页面坐标
                        min_x = left_bound * rect.width / width
                        min_y = top_bound * rect.height / height
                        max_x = right_bound * rect.width / width
                        max_y = bottom_bound * rect.height / height
                        
                        content_rect = fitz.Rect(min_x, min_y, max_x, max_y)
                        
                        # 创建可视化图像，在原始图像上绘制检测到的内容区域
                        if debug_dir and channels >= 3:
                            debug_visual_path = os.path.join(debug_dir, f"page_{page_num+1}_content_rect.png")
                            visual_img = np_img.copy()
                            # 绘制矩形边界
                            visual_img[top_bound:bottom_bound+1, left_bound:left_bound+5] = [255, 0, 0]  # 左边界
                            visual_img[top_bound:bottom_bound+1, right_bound-4:right_bound+1] = [255, 0, 0]  # 右边界
                            visual_img[top_bound:top_bound+5, left_bound:right_bound+1] = [255, 0, 0]  # 上边界
                            visual_img[bottom_bound-4:bottom_bound+1, left_bound:right_bound+1] = [255, 0, 0]  # 下边界
                            Image.fromarray(visual_img).save(debug_visual_path)
                    else:
                        content_rect = rect  # 未发现内容，使用整个页面
                except Exception as e:
                    print(f"像素分析出错: {str(e)}")
                    content_rect = rect  # 出错时使用整个页面
                
                # 内容区域有效性验证
                # 防止裁剪过多 - 如果内容区域太小，可能是错误检测
                # if content_rect.width < rect.width * 0.1 or content_rect.height < rect.height * 0.1:
                #     print(f"检测到的内容区域过小，使用整个页面: {content_rect}")
                #     content_rect = rect
                
                # 防止裁剪过少 - 如果内容区域几乎和页面一样大，微调一下裁剪区域
                # if content_rect.width > rect.width * 0.98 or content_rect.height > rect.height * 0.98:
                #     margin_x = rect.width * 0.02
                #     margin_y = rect.height * 0.02
                #     content_rect = fitz.Rect(margin_x, margin_y, rect.width - margin_x, rect.height - margin_y)
                
                # 应用边距
                crop_box = fitz.Rect(
                    max(content_rect.x0 - left_margin, 0),
                    max(content_rect.y0 - top_margin, 0),
                    min(content_rect.x1 + right_margin, rect.width),
                    min(content_rect.y1 + bottom_margin, rect.height)
                )
                
                # 确保裁剪框不超出页面边界
                crop_box = crop_box & rect
                
                # 创建新页面并插入裁剪后的内容
                new_page = new_doc.new_page(width=crop_box.width, height=crop_box.height)
                new_page.show_pdf_page(new_page.rect, doc, page_num, clip=crop_box)
                
            except Exception as e:
                # 如果处理当前页面出错，保留原始页面
                print(f"处理第 {page_num+1} 页时出错: {str(e)}")
                page = doc.load_page(page_num)
                new_page = new_doc.new_page(width=page.rect.width, height=page.rect.height)
                new_page.show_pdf_page(new_page.rect, doc, page_num)
        
        # 如果是覆盖原文件，先保存为临时文件，然后替换
        if input_path == output_path:
            temp_path = output_path + ".temp"
            new_doc.save(temp_path)
            new_doc.close()
            doc.close()
            os.replace(temp_path, output_path)
        else:
            # 直接保存到新位置
            new_doc.save(output_path)
            new_doc.close()
            doc.close()
    
    
    def on_frame_configure(self, event):
        """合并内容区布局更新，避免缩放时频繁重排。"""
        self.request_scroll_layout_update()
        
    def on_canvas_configure(self, event):
        """合并画布宽度更新，避免缩放时卡顿。"""
        self.request_scroll_layout_update(event.width)
    
    def delayed_layout_update(self):
        """手动刷新滚动区域。"""
        self.request_scroll_layout_update()

    def request_scroll_layout_update(self, width=None):
        if not hasattr(self, 'scroll_canvas'):
            return

        if width is not None:
            self._pending_canvas_width = width

        if self._layout_update_job is None:
            self._layout_update_job = self.root.after_idle(self.apply_scroll_layout_update)

    def apply_scroll_layout_update(self):
        self._layout_update_job = None

        if not hasattr(self, 'scroll_canvas'):
            return

        target_width = self._pending_canvas_width
        if target_width is not None and target_width != self._last_canvas_width:
            self.scroll_canvas.itemconfigure(self.canvas_window, width=target_width)
            self._last_canvas_width = target_width
        self._pending_canvas_width = None

        self.scroll_canvas.configure(scrollregion=self.scroll_canvas.bbox("all"))
        self.update_scrollbar_visibility()

    def update_scrollbar_visibility(self):
        if not hasattr(self, 'scroll_canvas'):
            return

        content_height = self.content_frame.winfo_reqheight()
        canvas_height = self.scroll_canvas.winfo_height()
        should_reset = content_height <= canvas_height + 4

        if should_reset:
            self.scroll_canvas.yview_moveto(0)

    def on_mousewheel(self, event):
        if not hasattr(self, 'scroll_canvas'):
            return

        content_height = self.content_frame.winfo_reqheight()
        canvas_height = self.scroll_canvas.winfo_height()
        if content_height <= canvas_height + 4:
            return

        if getattr(event, "num", None) == 4:
            delta = -1
        elif getattr(event, "num", None) == 5:
            delta = 1
        else:
            delta = -1 * int(event.delta / 120) if event.delta else 0

        if delta:
            self.scroll_canvas.yview_scroll(delta, "units")

def main():
    # 创建TkinterDnD应用
    enable_high_dpi()
    root = TkinterDnD.Tk()
    try:
        root.tk.call("tk", "scaling", root.winfo_fpixels("1i") / 72.0)
    except tk.TclError:
        pass
    app = PDFCropperApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
