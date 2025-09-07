"""
GUI样式配置模块
定义界面的颜色、字体、布局等样式
"""

import tkinter as tk
from tkinter import ttk


class AppStyles:
    """应用程序样式配置类"""
    
    # 颜色配置
    COLORS = {
        'primary': '#2E86AB',      # 主色调 - 蓝色
        'secondary': '#A23B72',    # 次要色 - 紫色
        'success': '#28A745',      # 成功色 - 绿色
        'warning': '#FFC107',      # 警告色 - 黄色
        'danger': '#DC3545',       # 危险色 - 红色
        'info': '#17A2B8',         # 信息色 - 青色
        'light': '#F8F9FA',        # 浅色
        'dark': '#343A40',         # 深色
        'white': '#FFFFFF',        # 白色
        'gray_100': '#F8F9FA',
        'gray_200': '#E9ECEF',
        'gray_300': '#DEE2E6',
        'gray_400': '#CED4DA',
        'gray_500': '#ADB5BD',
        'gray_600': '#6C757D',
        'gray_700': '#495057',
        'gray_800': '#343A40',
        'gray_900': '#212529',
    }
    
    # 字体配置
    FONTS = {
        'default': ('Microsoft YaHei UI', 9),
        'heading': ('Microsoft YaHei UI', 12, 'bold'),
        'subheading': ('Microsoft YaHei UI', 10, 'bold'),
        'small': ('Microsoft YaHei UI', 8),
        'monospace': ('Consolas', 9),
        'large': ('Microsoft YaHei UI', 11),
    }
    
    # 间距配置
    SPACING = {
        'xs': 2,
        'sm': 5,
        'md': 10,
        'lg': 15,
        'xl': 20,
        'xxl': 30,
    }
    
    @classmethod
    def configure_ttk_styles(cls, root):
        """配置ttk样式"""
        style = ttk.Style(root)
        
        # 设置主题
        try:
            style.theme_use('clam')  # 使用clam主题作为基础
        except:
            pass
        
        # 配置按钮样式
        style.configure('Primary.TButton',
                       background=cls.COLORS['primary'],
                       foreground=cls.COLORS['white'],
                       font=cls.FONTS['default'],
                       padding=(10, 5))
        
        style.map('Primary.TButton',
                 background=[('active', '#1e5f7a'),
                           ('pressed', '#1a5269')])
        
        style.configure('Success.TButton',
                       background=cls.COLORS['success'],
                       foreground=cls.COLORS['white'],
                       font=cls.FONTS['default'],
                       padding=(10, 5))
        
        style.map('Success.TButton',
                 background=[('active', '#1e7e34'),
                           ('pressed', '#1c7430')])
        
        style.configure('Warning.TButton',
                       background=cls.COLORS['warning'],
                       foreground=cls.COLORS['dark'],
                       font=cls.FONTS['default'],
                       padding=(10, 5))
        
        style.map('Warning.TButton',
                 background=[('active', '#e0a800'),
                           ('pressed', '#d39e00')])
        
        style.configure('Danger.TButton',
                       background=cls.COLORS['danger'],
                       foreground=cls.COLORS['white'],
                       font=cls.FONTS['default'],
                       padding=(10, 5))
        
        style.map('Danger.TButton',
                 background=[('active', '#c82333'),
                           ('pressed', '#bd2130')])
        
        # 配置标签框样式
        style.configure('Card.TLabelframe',
                       background=cls.COLORS['white'],
                       borderwidth=1,
                       relief='solid',
                       font=cls.FONTS['subheading'])
        
        style.configure('Card.TLabelframe.Label',
                       background=cls.COLORS['white'],
                       foreground=cls.COLORS['dark'],
                       font=cls.FONTS['subheading'])
        
        # 配置进度条样式
        style.configure('Success.Horizontal.TProgressbar',
                       background=cls.COLORS['success'],
                       troughcolor=cls.COLORS['gray_200'],
                       borderwidth=0,
                       lightcolor=cls.COLORS['success'],
                       darkcolor=cls.COLORS['success'])
        
        # 配置笔记本样式
        style.configure('TNotebook',
                       background=cls.COLORS['gray_100'],
                       borderwidth=0)
        
        style.configure('TNotebook.Tab',
                       background=cls.COLORS['gray_200'],
                       foreground=cls.COLORS['dark'],
                       font=cls.FONTS['default'],
                       padding=(12, 8))
        
        style.map('TNotebook.Tab',
                 background=[('selected', cls.COLORS['white']),
                           ('active', cls.COLORS['gray_300'])],
                 foreground=[('selected', cls.COLORS['primary'])])
        
        # 配置树形视图样式
        style.configure('Treeview',
                       background=cls.COLORS['white'],
                       foreground=cls.COLORS['dark'],
                       font=cls.FONTS['default'],
                       fieldbackground=cls.COLORS['white'])
        
        style.configure('Treeview.Heading',
                       background=cls.COLORS['gray_200'],
                       foreground=cls.COLORS['dark'],
                       font=cls.FONTS['subheading'])
        
        style.map('Treeview',
                 background=[('selected', cls.COLORS['primary'])],
                 foreground=[('selected', cls.COLORS['white'])])
        
        # 配置分隔符样式
        style.configure('TSeparator',
                       background=cls.COLORS['gray_300'])
        
        return style
    
    @classmethod
    def configure_text_widget(cls, text_widget, style_type='default'):
        """配置文本组件样式"""
        if style_type == 'log':
            text_widget.configure(
                background=cls.COLORS['gray_900'],
                foreground=cls.COLORS['gray_100'],
                font=cls.FONTS['monospace'],
                insertbackground=cls.COLORS['white'],
                selectbackground=cls.COLORS['primary'],
                selectforeground=cls.COLORS['white'],
                wrap=tk.WORD,
                padx=cls.SPACING['sm'],
                pady=cls.SPACING['sm']
            )
        elif style_type == 'stats':
            text_widget.configure(
                background=cls.COLORS['white'],
                foreground=cls.COLORS['dark'],
                font=cls.FONTS['default'],
                insertbackground=cls.COLORS['dark'],
                selectbackground=cls.COLORS['primary'],
                selectforeground=cls.COLORS['white'],
                wrap=tk.WORD,
                padx=cls.SPACING['sm'],
                pady=cls.SPACING['sm'],
                relief='flat',
                borderwidth=0
            )
        else:
            text_widget.configure(
                background=cls.COLORS['white'],
                foreground=cls.COLORS['dark'],
                font=cls.FONTS['default'],
                insertbackground=cls.COLORS['dark'],
                selectbackground=cls.COLORS['primary'],
                selectforeground=cls.COLORS['white']
            )
    
    @classmethod
    def get_icon_color(cls, icon_type):
        """获取图标颜色"""
        icon_colors = {
            'success': cls.COLORS['success'],
            'warning': cls.COLORS['warning'],
            'error': cls.COLORS['danger'],
            'info': cls.COLORS['info'],
            'primary': cls.COLORS['primary'],
            'secondary': cls.COLORS['secondary'],
        }
        return icon_colors.get(icon_type, cls.COLORS['dark'])
    
    @classmethod
    def create_tooltip_style(cls):
        """创建工具提示样式"""
        return {
            'background': cls.COLORS['gray_800'],
            'foreground': cls.COLORS['white'],
            'font': cls.FONTS['small'],
            'relief': 'solid',
            'borderwidth': 1,
            'padx': cls.SPACING['sm'],
            'pady': cls.SPACING['xs']
        }


class ToolTip:
    """工具提示类"""
    
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        self.widget.bind('<Enter>', self.on_enter)
        self.widget.bind('<Leave>', self.on_leave)
        self.widget.bind('<Motion>', self.on_motion)
    
    def on_enter(self, event=None):
        """鼠标进入时显示提示"""
        self.show_tooltip()
    
    def on_leave(self, event=None):
        """鼠标离开时隐藏提示"""
        self.hide_tooltip()
    
    def on_motion(self, event=None):
        """鼠标移动时更新提示位置"""
        if self.tooltip_window:
            self.update_position(event)
    
    def show_tooltip(self):
        """显示工具提示"""
        if self.tooltip_window or not self.text:
            return
        
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + 20
        
        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{x}+{y}")
        
        style_config = AppStyles.create_tooltip_style()
        label = tk.Label(self.tooltip_window, text=self.text, **style_config)
        label.pack()
    
    def hide_tooltip(self):
        """隐藏工具提示"""
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None
    
    def update_position(self, event):
        """更新提示位置"""
        if self.tooltip_window:
            x = self.widget.winfo_rootx() + event.x + 10
            y = self.widget.winfo_rooty() + event.y + 10
            self.tooltip_window.wm_geometry(f"+{x}+{y}")


def apply_modern_style(root):
    """应用现代化样式到根窗口"""
    # 配置根窗口
    root.configure(bg=AppStyles.COLORS['gray_100'])
    
    # 配置ttk样式
    style = AppStyles.configure_ttk_styles(root)
    
    return style


def create_styled_button(parent, text, command=None, style_type='primary', tooltip=None):
    """创建带样式的按钮"""
    style_map = {
        'primary': 'Primary.TButton',
        'success': 'Success.TButton',
        'warning': 'Warning.TButton',
        'danger': 'Danger.TButton',
    }
    
    button = ttk.Button(parent, text=text, command=command, 
                       style=style_map.get(style_type, 'Primary.TButton'))
    
    if tooltip:
        ToolTip(button, tooltip)
    
    return button


def create_card_frame(parent, title, padding=10):
    """创建卡片样式的框架"""
    frame = ttk.LabelFrame(parent, text=title, style='Card.TLabelframe', padding=padding)
    return frame 