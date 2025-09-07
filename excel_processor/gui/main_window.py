import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
from pathlib import Path
import sys
from pathlib import Path

# 添加项目根目录到Python路径
current_dir = Path(__file__).parent.parent
sys.path.insert(0, str(current_dir))

from gui.styles import AppStyles, apply_modern_style, create_styled_button, create_card_frame, ToolTip


class MainWindow:
    """主窗口类，实现现代化GUI界面"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.config = self.load_config()
        self.setup_window()
        self.create_widgets()
        self.setup_layout()
        
    def load_config(self):
        """加载配置文件"""
        try:
            config_path = Path(__file__).parent.parent / "resources" / "config.json"
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"加载配置文件失败: {e}")
            return {
                "window": {"width": 1200, "height": 800, "title": "Excel IP地址处理工具 v2.0"},
                "ui": {"theme": "default", "font_size": 10}
            }
    
    def setup_window(self):
        """设置主窗口属性"""
        window_config = self.config.get("window", {})
        
        # 设置窗口标题和大小
        self.root.title(window_config.get("title", "Excel IP地址处理工具 v2.0"))
        self.root.geometry(f"{window_config.get('width', 1200)}x{window_config.get('height', 800)}")
        self.root.minsize(window_config.get('min_width', 800), window_config.get('min_height', 600))
        
        # 应用现代化样式
        self.style = apply_modern_style(self.root)
        
        # 设置窗口居中
        self.center_window()
        
        # 设置窗口图标（如果有的话）
        try:
            # self.root.iconbitmap("icon.ico")  # 可以后续添加图标
            pass
        except:
            pass
    
    def center_window(self):
        """将窗口居中显示"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
    
    def create_widgets(self):
        """创建所有GUI组件"""
        # 创建主框架
        self.main_frame = ttk.Frame(self.root)
        
        # 创建菜单栏
        self.create_menu()
        
        # 创建工具栏
        self.create_toolbar()
        
        # 创建左侧文件操作区
        self.create_file_operations_panel()
        
        # 创建右侧主显示区
        self.create_main_display_panel()
        
        # 创建状态栏
        self.create_status_bar()
    
    def create_menu(self):
        """创建菜单栏"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 文件菜单
        self.file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件", menu=self.file_menu)
        self.file_menu.add_command(label="打开文件...", command=self.select_file, accelerator="Ctrl+O")
        self.file_menu.add_separator()
        self.file_menu.add_command(label="退出", command=self.root.quit, accelerator="Ctrl+Q")
        
        # 工具菜单
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="工具", menu=tools_menu)
        tools_menu.add_command(label="清空日志", command=self.clear_log)
        tools_menu.add_command(label="清空历史", command=self.clear_history)
        
        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="使用说明", command=self.show_help)
        help_menu.add_command(label="关于", command=self.show_about)
        
        # 绑定快捷键
        self.root.bind('<Control-o>', lambda e: self.select_file())
        self.root.bind('<Control-q>', lambda e: self.root.quit())
    
    def create_toolbar(self):
        """创建工具栏"""
        self.toolbar = ttk.Frame(self.root)
        
        # 文件选择按钮
        self.select_btn = ttk.Button(self.toolbar, text="选择文件", command=self.select_file)
        self.select_btn.pack(side=tk.LEFT, padx=5, pady=5)
        
        # 处理按钮
        self.process_btn = ttk.Button(self.toolbar, text="开始处理", command=self.start_processing, state=tk.DISABLED)
        self.process_btn.pack(side=tk.LEFT, padx=5, pady=5)
        
        # 取消按钮
        self.cancel_btn = ttk.Button(self.toolbar, text="取消处理", command=self.cancel_processing, state=tk.DISABLED)
        self.cancel_btn.pack(side=tk.LEFT, padx=5, pady=5)
        
        # 分隔符
        ttk.Separator(self.toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=5)
        
        # 清空日志按钮
        self.clear_log_btn = ttk.Button(self.toolbar, text="清空日志", command=self.clear_log)
        self.clear_log_btn.pack(side=tk.LEFT, padx=5, pady=5)
    
    def create_file_operations_panel(self):
        """创建左侧文件操作面板"""
        self.left_panel = create_card_frame(self.main_frame, "文件操作", padding=10)
        
        # 当前文件显示
        ttk.Label(self.left_panel, text="当前文件:").pack(anchor=tk.W, pady=(0, 5))
        self.file_var = tk.StringVar(value="未选择文件")
        self.file_label = ttk.Label(self.left_panel, textvariable=self.file_var, 
                                   foreground="blue", cursor="hand2")
        self.file_label.pack(anchor=tk.W, pady=(0, 10))
        
        # 处理选项
        ttk.Label(self.left_panel, text="处理选项:").pack(anchor=tk.W, pady=(10, 5))
        
        self.auto_backup_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self.left_panel, text="自动备份原文件", 
                       variable=self.auto_backup_var).pack(anchor=tk.W, pady=2)
        
        self.show_progress_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(self.left_panel, text="显示详细进度", 
                       variable=self.show_progress_var).pack(anchor=tk.W, pady=2)
        
        # 进度显示
        ttk.Label(self.left_panel, text="处理进度:").pack(anchor=tk.W, pady=(20, 5))
        self.progress_var = tk.StringVar(value="就绪")
        ttk.Label(self.left_panel, textvariable=self.progress_var).pack(anchor=tk.W, pady=(0, 5))
        
        self.progress_bar = ttk.Progressbar(self.left_panel, mode='determinate', style='Success.Horizontal.TProgressbar')
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))
        
        # 统计信息
        self.stats_frame = create_card_frame(self.left_panel, "处理统计", padding=5)
        self.stats_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.stats_text = tk.Text(self.stats_frame, height=8, width=30, state=tk.DISABLED)
        AppStyles.configure_text_widget(self.stats_text, 'stats')
        self.stats_text.pack(fill=tk.BOTH, expand=True)
    
    def create_main_display_panel(self):
        """创建右侧主显示面板"""
        self.right_panel = ttk.Frame(self.main_frame)
        
        # 创建笔记本控件（标签页）
        self.notebook = ttk.Notebook(self.right_panel)
        
        # 日志显示标签页
        self.log_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.log_frame, text="处理日志")
        
        # 创建日志显示区域
        log_container = ttk.Frame(self.log_frame)
        log_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.log_text = tk.Text(log_container, wrap=tk.WORD, state=tk.DISABLED)
        AppStyles.configure_text_widget(self.log_text, 'log')
        log_scrollbar = ttk.Scrollbar(log_container, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 历史记录标签页
        self.history_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.history_frame, text="历史记录")
        
        # 创建历史记录显示区域
        history_container = ttk.Frame(self.history_frame)
        history_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 历史记录树形视图
        self.history_tree = ttk.Treeview(history_container, columns=('file', 'time', 'status', 'results'), show='headings')
        self.history_tree.heading('file', text='文件名')
        self.history_tree.heading('time', text='处理时间')
        self.history_tree.heading('status', text='状态')
        self.history_tree.heading('results', text='处理结果')
        
        history_scrollbar = ttk.Scrollbar(history_container, orient=tk.VERTICAL, command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=history_scrollbar.set)
        
        self.history_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        history_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.notebook.pack(fill=tk.BOTH, expand=True)
    
    def create_status_bar(self):
        """创建状态栏"""
        self.status_bar = ttk.Frame(self.root)
        
        self.status_var = tk.StringVar(value="就绪")
        self.status_label = ttk.Label(self.status_bar, textvariable=self.status_var)
        self.status_label.pack(side=tk.LEFT, padx=5)
        
        # 添加分隔符
        ttk.Separator(self.status_bar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        # 显示当前时间
        self.time_var = tk.StringVar()
        self.time_label = ttk.Label(self.status_bar, textvariable=self.time_var)
        self.time_label.pack(side=tk.RIGHT, padx=5)
        
        self.update_time()
    
    def setup_layout(self):
        """设置布局"""
        # 工具栏
        self.toolbar.pack(fill=tk.X, padx=5, pady=5)
        
        # 主框架
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 左右面板
        self.left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 5))
        self.right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 状态栏
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM)
    
    def update_time(self):
        """更新状态栏时间显示"""
        import datetime
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.time_var.set(current_time)
        self.root.after(1000, self.update_time)
    
    # 事件处理方法（占位符，后续实现）
    def select_file(self):
        """选择文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[
                ("Excel文件", "*.xlsx *.xls"),
                ("Excel 2007-365", "*.xlsx"),
                ("Excel 97-2003", "*.xls"),
                ("所有文件", "*.*")
            ]
        )
        
        if file_path:
            self.file_var.set(os.path.basename(file_path))
            self.current_file = file_path
            self.process_btn.config(state=tk.NORMAL)
            self.log_message(f"已选择文件: {file_path}")
            self.status_var.set(f"已选择文件: {os.path.basename(file_path)}")
    
    def start_processing(self):
        """开始处理"""
        self.log_message("开始处理文件...")
        self.status_var.set("正在处理...")
        # TODO: 实现实际的处理逻辑
    
    def cancel_processing(self):
        """取消处理"""
        self.log_message("用户取消了处理操作")
        self.status_var.set("已取消")
    
    def clear_log(self):
        """清空日志"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.status_var.set("日志已清空")
    
    def clear_history(self):
        """清空历史记录"""
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        self.status_var.set("历史记录已清空")
    
    def show_help(self):
        """显示帮助"""
        help_text = """
Excel IP地址处理工具 v2.0 使用说明

主要功能：
1. IP地址格式标准化
2. 重复IP检测和处理
3. 网段包含关系检测
4. 网段一致性检查

使用步骤：
1. 点击"选择文件"或使用Ctrl+O选择Excel文件
2. 配置处理选项
3. 点击"开始处理"
4. 查看处理日志和结果统计
5. 处理完成后可在历史记录中查看详情

快捷键：
- Ctrl+O: 打开文件
- Ctrl+Q: 退出程序
        """
        messagebox.showinfo("使用说明", help_text)
    
    def show_about(self):
        """显示关于信息"""
        about_text = """
Excel IP地址处理工具 v2.0

版本: 2.0.0
作者: Excel Processor Team

这是一个专业的Excel IP地址处理工具，
支持IP地址标准化、重复检测、网段分析等功能。

© 2024 All Rights Reserved
        """
        messagebox.showinfo("关于", about_text)
    
    def update_file_menu_command(self, new_command):
        """更新文件菜单中的打开文件命令"""
        try:
            self.file_menu.entryconfig(0, command=new_command)
        except Exception as e:
            print(f"更新菜单命令失败: {e}")
    
    def log_message(self, message):
        """添加日志消息"""
        import datetime
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"
        
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, formatted_message)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
    
    def run(self):
        """运行主窗口"""
        self.log_message("Excel IP地址处理工具已启动")
        self.root.mainloop()


if __name__ == "__main__":
    app = MainWindow()
    app.run() 