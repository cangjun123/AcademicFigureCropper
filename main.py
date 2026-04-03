import os
import sys
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

class PDFCropperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Academic Figure Cropper")
        self.root.geometry("500x700")  # 增加默认窗口大小
        self.root.minsize(500, 700)  # 增加最小窗口大小
        
        # 设置窗口图标（如果有的话）
        try:
            icon_path = resource_path("icon.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"加载图标失败: {e}")
            pass
        
        # 设置窗口始终置顶
        self.root.attributes("-topmost", True)
        
        # 设置主题色 - 更现代的配色方案
        self.primary_color = "#2979ff"  # 蓝色主色调
        self.accent_color = "#ff9100"   # 橙色强调色
        self.bg_color = "#fafafa"       # 浅灰背景
        self.card_bg_color = "#ffffff"  # 卡片背景色
        self.text_color = "#212121"     # 深灰文本
        self.secondary_text = "#757575" # 次要文本颜色
        self.button_text_color = "#ffffff"  # 按钮文字颜色为白色
        self.border_color = "#e0e0e0"   # 边框颜色
        
        # 支持的图片格式
        self.supported_img_formats = ('.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.gif')
        self.save_debug_images = False
        
        # 配置样式
        self.style = ttk.Style()
        self.style.configure("TFrame", background=self.bg_color)
        self.style.configure("Card.TFrame", background=self.card_bg_color, relief="flat")
        
        # 标签样式
        self.style.configure("TLabel", background=self.bg_color, foreground=self.text_color, font=("微软雅黑", 9))
        self.style.configure("Title.TLabel", background=self.bg_color, foreground=self.primary_color, font=("微软雅黑", 18, "bold"))
        self.style.configure("Subtitle.TLabel", background=self.bg_color, foreground=self.secondary_text, font=("微软雅黑", 11))
        self.style.configure("Card.TLabel", background=self.card_bg_color, foreground=self.text_color, font=("微软雅黑", 9))
        
        # 按钮样式
        self.style.configure("TButton", background=self.primary_color, foreground=self.button_text_color, font=("微软雅黑", 9))
        self.style.map("TButton", 
                      background=[('active', self.primary_color), ('pressed', '#1c54b2')],
                      foreground=[('active', self.button_text_color), ('pressed', self.button_text_color)])
        
        self.style.configure("Accent.TButton", background=self.accent_color, foreground=self.button_text_color, font=("微软雅黑", 9))
        self.style.map("Accent.TButton", 
                      background=[('active', self.accent_color), ('pressed', '#c56200')],
                      foreground=[('active', self.button_text_color), ('pressed', self.button_text_color)])
        
        # 复选框样式
        self.style.configure("TCheckbutton", background=self.bg_color, font=("微软雅黑", 9))
        self.style.map("TCheckbutton", 
                     background=[('active', self.bg_color)],
                     foreground=[('active', self.primary_color)])
        
        # 框架样式
        self.style.configure("TLabelframe", background=self.card_bg_color, bordercolor=self.border_color)
        self.style.configure("TLabelframe.Label", background=self.bg_color, foreground=self.primary_color, font=("微软雅黑", 10, "bold"))
        
        # 滚动条样式
        self.style.configure("Vertical.TScrollbar", gripcount=0, background=self.bg_color, troughcolor=self.bg_color, 
                          arrowcolor=self.primary_color, bordercolor=self.border_color)
        
        # 进度条样式
        self.style.configure("TProgressbar", background=self.primary_color, troughcolor=self.border_color, 
                          bordercolor=self.border_color, lightcolor=self.primary_color, darkcolor=self.primary_color)
        
        # 设置背景色
        self.root.configure(bg=self.bg_color)
        
        # 创建配置文件管理
        config_name = "pdf_cropper_config.ini"
        # 使用用户目录保存配置文件，确保打包后也能正常读写
        user_dir = os.path.expanduser("~")
        self.config_file = os.path.join(user_dir, config_name)
        self.config = configparser.ConfigParser()
        self.load_config()
        
        # 进度变量
        self.progress_var = tk.DoubleVar()
        self.progress_var.set(0)
        self.ui_queue = queue.Queue()
        
        # 创建UI元素
        self.create_widgets()
        
        # 处理的文件列表
        self.processing_files = []
        
        # 在初始化完成后进行更新，确保组件正确显示
        self.root.update_idletasks()
        self.on_frame_configure(None)
        
        # 设置定时器更新界面布局，确保滚动区域计算准确
        self.root.after(100, self.delayed_layout_update)
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
    
    def save_config(self):
        """保存配置到文件"""
        with open(self.config_file, 'w') as f:
            self.config.write(f)
    
    def create_widgets(self):
        """创建UI元素"""
        # 创建底部状态栏（先创建，确保它在底部）
        self.status_bar = ttk.Frame(self.root, relief=tk.SUNKEN, style="TFrame")
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.status_text = ttk.Label(self.status_bar, text="就绪", padding=(5, 2))
        self.status_text.pack(side=tk.LEFT)
        
        self.version_label = ttk.Label(self.status_bar, text="v1.1", padding=(5, 2))
        self.version_label.pack(side=tk.RIGHT)
        
        # 创建带滚动条的主容器
        self.canvas = tk.Canvas(self.root, bg=self.bg_color, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview, style="Vertical.TScrollbar")
        
        # 配置布局
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 关联滚动条和画布
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # 创建主框架
        self.main_frame = ttk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window((0, 0), window=self.main_frame, anchor="nw", tags="self.main_frame")
        
        # 绑定事件处理滚动和调整大小
        self.main_frame.bind("<Configure>", self.on_frame_configure)
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        
        # 添加边距
        self.main_frame.pack_configure(padx=20, pady=20)
        
        # 创建标题和描述区域
        title_frame = ttk.Frame(self.main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 20))
        
        title_label = ttk.Label(title_frame, text="Academic Figure Cropper", style="Title.TLabel")
        title_label.pack(side=tk.TOP, anchor=tk.W)
        
        description = ttk.Label(title_frame, 
                              text="拖放PDF或图片文件到下方区域，自动剪裁白边并保存", 
                              style="Subtitle.TLabel")
        description.pack(side=tk.TOP, anchor=tk.W, pady=(5, 0))
        
        # 拖放区域 - 使用卡片风格
        self.drop_frame = ttk.LabelFrame(self.main_frame, text="拖放区域")
        self.drop_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        self.drop_area = ttk.Frame(self.drop_frame, style="Card.TFrame")
        self.drop_area.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # 创建一个居中的容器
        drop_center_frame = ttk.Frame(self.drop_area, style="Card.TFrame")
        drop_center_frame.pack(expand=True, pady=20)
        
        self.drop_icon_label = ttk.Label(drop_center_frame, text="📄", font=("微软雅黑", 48), style="Card.TLabel")
        self.drop_icon_label.pack(pady=(10, 15))
        
        self.drop_label = ttk.Label(drop_center_frame, text="拖放PDF或图片文件到这里", font=("微软雅黑", 12, "bold"), style="Card.TLabel")
        self.drop_label.pack(pady=(0, 10))
        
        self.drop_hint = ttk.Label(drop_center_frame, text="支持PDF和常见图片格式(.jpg, .png等)", style="Card.TLabel")
        self.drop_hint.pack()
        
        # 为整个拖放区域绑定拖放事件
        for widget in [self.drop_area, drop_center_frame, self.drop_icon_label, self.drop_label, self.drop_hint]:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind('<<Drop>>', self.drop)
        
        # 进度条和状态
        progress_frame = ttk.Frame(self.main_frame, style="TFrame")
        progress_frame.pack(fill=tk.X, pady=(0, 20))
        
        self.status_var = tk.StringVar()
        self.status_var.set("准备就绪")
        
        self.status_label = ttk.Label(progress_frame, textvariable=self.status_var)
        self.status_label.pack(side=tk.TOP, anchor=tk.W, pady=(0, 5))
        
        self.progress = ttk.Progressbar(progress_frame, variable=self.progress_var, length=400, mode='determinate', style="TProgressbar")
        self.progress.pack(fill=tk.X)
        
        # 设置区域 - 使用卡片风格
        self.settings_frame = ttk.LabelFrame(self.main_frame, text="设置")
        self.settings_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 设置内容框架
        settings_content = ttk.Frame(self.settings_frame, style="Card.TFrame")
        settings_content.pack(fill=tk.X, padx=15, pady=15)
        
        # 文件输出设置
        output_frame = ttk.Frame(settings_content, style="Card.TFrame")
        output_frame.pack(fill=tk.X, pady=(0, 15))
        
        output_label = ttk.Label(output_frame, text="文件输出", font=("微软雅黑", 11, "bold"), foreground=self.primary_color, style="Card.TLabel")
        output_label.grid(row=0, column=0, sticky=tk.W, pady=5)
        
        self.overwrite_var = tk.BooleanVar()
        self.overwrite_var.set(self.config.getboolean('Settings', 'overwrite_original'))
        self.overwrite_check = ttk.Checkbutton(output_frame, text="覆盖原文件", 
                                             variable=self.overwrite_var, 
                                             command=self.toggle_output_path,
                                             style="TCheckbutton")
        self.overwrite_check.grid(row=1, column=0, sticky=tk.W, padx=(20, 0))
        
        output_dir_frame = ttk.Frame(output_frame, style="Card.TFrame")
        output_dir_frame.grid(row=2, column=0, sticky=tk.W, padx=(20, 0), pady=5)
        
        self.output_path_var = tk.StringVar()
        saved_path = self.config.get('Settings', 'output_dir')
        self.output_path_var.set(saved_path if saved_path else "未设置")
        
        self.output_entry = ttk.Entry(output_dir_frame, textvariable=self.output_path_var, width=35)
        self.output_entry.pack(side=tk.LEFT, padx=(0, 5))
        
        self.output_button = ttk.Button(output_dir_frame, text="浏览...", 
                                      command=self.select_output_dir,
                                      style="TButton",
                                      state=tk.NORMAL if not self.overwrite_var.get() else tk.DISABLED)
        self.output_button.pack(side=tk.LEFT)
        
        # 边距设置框架
        margins_frame = ttk.Frame(settings_content, style="Card.TFrame")
        margins_frame.pack(fill=tk.X, pady=5)
        
        margins_label = ttk.Label(margins_frame, text="剪裁边距", font=("微软雅黑", 11, "bold"), foreground=self.primary_color, style="Card.TLabel")
        margins_label.grid(row=0, column=0, columnspan=4, sticky=tk.W, pady=5)
        
        # 左边距
        left_frame = ttk.Frame(margins_frame, style="Card.TFrame")
        left_frame.grid(row=1, column=0, padx=(20, 15), pady=10, sticky=tk.W)
        
        ttk.Label(left_frame, text="左边距:", style="Card.TLabel").pack(side=tk.LEFT)
        self.left_margin_var = tk.IntVar(value=int(self.config.get('Settings', 'left_margin')))
        self.left_margin_spin = ttk.Spinbox(left_frame, from_=0, to=50, width=5, 
                                          textvariable=self.left_margin_var,
                                          command=self.save_margins)
        self.left_margin_spin.pack(side=tk.LEFT, padx=(5, 0))
        
        # 右边距
        right_frame = ttk.Frame(margins_frame, style="Card.TFrame")
        right_frame.grid(row=1, column=1, padx=15, pady=10, sticky=tk.W)
        
        ttk.Label(right_frame, text="右边距:", style="Card.TLabel").pack(side=tk.LEFT)
        self.right_margin_var = tk.IntVar(value=int(self.config.get('Settings', 'right_margin')))
        self.right_margin_spin = ttk.Spinbox(right_frame, from_=0, to=50, width=5, 
                                           textvariable=self.right_margin_var,
                                           command=self.save_margins)
        self.right_margin_spin.pack(side=tk.LEFT, padx=(5, 0))
        
        # 上边距
        top_frame = ttk.Frame(margins_frame, style="Card.TFrame")
        top_frame.grid(row=2, column=0, padx=(20, 15), pady=10, sticky=tk.W)
        
        ttk.Label(top_frame, text="上边距:", style="Card.TLabel").pack(side=tk.LEFT)
        self.top_margin_var = tk.IntVar(value=int(self.config.get('Settings', 'top_margin')))
        self.top_margin_spin = ttk.Spinbox(top_frame, from_=0, to=50, width=5, 
                                         textvariable=self.top_margin_var,
                                         command=self.save_margins)
        self.top_margin_spin.pack(side=tk.LEFT, padx=(5, 0))
        
        # 下边距
        bottom_frame = ttk.Frame(margins_frame, style="Card.TFrame")
        bottom_frame.grid(row=2, column=1, padx=15, pady=10, sticky=tk.W)
        
        ttk.Label(bottom_frame, text="下边距:", style="Card.TLabel").pack(side=tk.LEFT)
        self.bottom_margin_var = tk.IntVar(value=int(self.config.get('Settings', 'bottom_margin')))
        self.bottom_margin_spin = ttk.Spinbox(bottom_frame, from_=0, to=50, width=5, 
                                            textvariable=self.bottom_margin_var,
                                            command=self.save_margins)
        self.bottom_margin_spin.pack(side=tk.LEFT, padx=(5, 0))
    
    def toggle_output_path(self):
        """根据覆盖选项切换输出路径按钮状态"""
        if self.overwrite_var.get():
            self.output_button.config(state=tk.DISABLED)
            self.output_entry.config(state=tk.DISABLED)
        else:
            self.output_button.config(state=tk.NORMAL)
            self.output_entry.config(state=tk.NORMAL)
        
        # 更新配置
        self.config['Settings']['overwrite_original'] = str(self.overwrite_var.get())
        self.save_config()
    
    def save_margins(self):
        """保存边距设置"""
        self.config['Settings']['left_margin'] = str(self.left_margin_var.get())
        self.config['Settings']['right_margin'] = str(self.right_margin_var.get())
        self.config['Settings']['top_margin'] = str(self.top_margin_var.get())
        self.config['Settings']['bottom_margin'] = str(self.bottom_margin_var.get())
        self.save_config()
    
    def select_output_dir(self):
        """选择输出目录"""
        output_dir = filedialog.askdirectory()
        if output_dir:
            self.output_path_var.set(output_dir)
            self.config['Settings']['output_dir'] = output_dir
            self.save_config()

    def get_processing_settings(self):
        """Capture UI-controlled settings before starting a worker thread."""
        return {
            'overwrite_original': self.overwrite_var.get(),
            'output_dir': self.config.get('Settings', 'output_dir').strip(),
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

    def finish_processing(self, total_success, total_failed):
        final_icon = "✅" if total_failed == 0 else "⚠️"
        self.drop_icon_label.config(text=final_icon)

        if total_failed > 0:
            self.status_var.set(f"处理完成: {total_success} 成功, {total_failed} 失败")
            messagebox.showinfo("完成", f"处理完成\n成功: {total_success} 个文件\n失败: {total_failed} 个文件")
        else:
            self.status_var.set(f"成功处理 {total_success} 个文件")
            messagebox.showinfo("完成", f"成功处理 {total_success} 个文件")

        self.root.after(1000, lambda: self.drop_icon_label.config(text="📄"))
    
    def drop(self, event):
        """处理文件拖放事件"""
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
        if not files:
            messagebox.showinfo("提示", "没有检测到支持的文件格式")
            return

        settings = self.get_processing_settings()

        # 检查输出路径
        if not settings['overwrite_original'] and not settings['output_dir']:
            messagebox.showwarning("警告", "请先选择输出文件夹")
            return

        if not settings['overwrite_original']:
            try:
                os.makedirs(settings['output_dir'], exist_ok=True)
            except OSError as exc:
                messagebox.showerror("错误", f"无法创建输出文件夹:\n{exc}")
                return

        self.progress_var.set(0)
        self.progress.config(maximum=len(files))

        # 更新UI反馈
        self.drop_icon_label.config(text="⏳")
        self.status_var.set("开始处理...")

        # 开始处理线程
        threading.Thread(target=self.process_files_thread, args=(files, settings), daemon=True).start()

    def process_files_thread(self, files, settings):
        """在单独的线程中处理文件"""
        total_success = 0
        total_failed = 0
        reserved_output_paths = set()

        for i, file_path in enumerate(files):
            try:
                filename = os.path.basename(file_path)
                self.enqueue_ui_call(self.status_var.set, f"正在处理: {filename}")

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
                self.enqueue_ui_call(
                    messagebox.showerror,
                    "错误",
                    f"处理文件 {os.path.basename(file_path)} 时出错:\n{str(e)}",
                )

        self.enqueue_ui_call(self.finish_processing, total_success, total_failed)

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
        """当框架大小改变时，更新滚动区域"""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
        # 当框架配置变化时，检查并更新滚动条状态
        if self.main_frame.winfo_reqheight() > self.canvas.winfo_height():
            self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)  # 显示滚动条
        else:
            self.scrollbar.pack_forget()  # 隐藏滚动条
        
    def on_canvas_configure(self, event):
        """当画布大小改变时，调整内部窗口大小"""
        width = event.width
        self.canvas.itemconfig(self.canvas_window, width=width)
        
        # 当画布大小改变时，检查并更新滚动条状态
        if self.main_frame.winfo_reqheight() > event.height:
            self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)  # 显示滚动条
        else:
            self.scrollbar.pack_forget()  # 隐藏滚动条
    
    def delayed_layout_update(self):
        """设置定时器更新界面布局，确保滚动区域计算准确"""
        self.on_frame_configure(None)
        # 只执行一次，不再继续调用

def main():
    # 创建TkinterDnD应用
    root = TkinterDnD.Tk()
    app = PDFCropperApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
