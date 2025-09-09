import sys
import io
from datetime import datetime
import logging


class GUILogHandler:
    """GUI日志处理器，将print输出重定向到GUI界面"""
    
    def __init__(self, log_widget=None):
        self.log_widget = log_widget
        self.current_log_level = LogLevel.INFO  # 默认日志级别
        # 在打包环境中，sys.stdout可能为None，需要处理这种情况
        self.original_stdout = sys.stdout if sys.stdout is not None else sys.__stdout__
        self.original_stderr = sys.stderr if sys.stderr is not None else sys.__stderr__
        
        # 如果仍然为None，创建一个虚拟的输出对象
        if self.original_stdout is None:
            self.original_stdout = io.StringIO()
        if self.original_stderr is None:
            self.original_stderr = io.StringIO()
            
        self.log_buffer = io.StringIO()
        
    def write(self, text):
        """重定向的写入方法"""
        # 安全地写入到原始输出（保持控制台输出）
        try:
            if self.original_stdout and hasattr(self.original_stdout, 'write'):
                self.original_stdout.write(text)
                if hasattr(self.original_stdout, 'flush'):
                    self.original_stdout.flush()
        except (AttributeError, OSError):
            # 在打包环境中可能出现写入错误，忽略
            pass
        
        # 如果有GUI组件，也写入到GUI
        if self.log_widget and text.strip():
            # 尝试从消息中提取日志级别
            level = self.extract_log_level(text.strip())
            self.write_to_gui(text.strip(), level)
    
    def flush(self):
        """刷新缓冲区"""
        try:
            if self.original_stdout and hasattr(self.original_stdout, 'flush'):
                self.original_stdout.flush()
        except (AttributeError, OSError):
            # 在打包环境中可能出现刷新错误，忽略
            pass
    
    def write_to_gui(self, message, level=None):
        """将消息写入GUI日志组件"""
        # 如果指定了级别，检查是否应该显示
        if level and not self.should_log(level):
            return
            
        if self.log_widget:
            try:
                timestamp = datetime.now().strftime("%H:%M:%S")
                
                # 根据日志级别设置颜色标签
                color_tag = self.get_color_tag(level) if level else ""
                formatted_message = f"[{timestamp}] {message}\n"
                
                # 启用文本框编辑
                self.log_widget.config(state='normal')
                
                # 插入新消息，如果有颜色标签则应用
                start_pos = self.log_widget.index('end')
                self.log_widget.insert('end', formatted_message)
                if color_tag:
                    end_pos = self.log_widget.index('end-1c')
                    self.log_widget.tag_add(color_tag, start_pos, end_pos)
                
                # 自动滚动到底部
                self.log_widget.see('end')
                
                # 禁用文本框编辑
                self.log_widget.config(state='disabled')
                
                # 更新界面
                self.log_widget.update_idletasks()
                
            except Exception as e:
                # 如果GUI写入失败，至少保证控制台输出
                self.original_stdout.write(f"GUI日志写入失败: {e}\n")
    
    def get_color_tag(self, level):
        """根据日志级别获取颜色标签"""
        color_map = {
            LogLevel.DEBUG: "debug",
            LogLevel.INFO: "info", 
            LogLevel.WARNING: "warning",
            LogLevel.ERROR: "error",
            LogLevel.CRITICAL: "critical"
        }
        return color_map.get(level, "")
    
    def configure_text_colors(self, text_widget):
        """配置文本组件的颜色标签"""
        text_widget.tag_configure("debug", foreground="#888888")
        text_widget.tag_configure("info", foreground="#E0E0E0")
        text_widget.tag_configure("warning", foreground="#FFD700")
        text_widget.tag_configure("error", foreground="#FF6B6B")
        text_widget.tag_configure("critical", foreground="#FFFFFF", background="#FF4444")
    
    def redirect_output(self):
        """开始重定向输出"""
        sys.stdout = self
        sys.stderr = self
    
    def restore_output(self):
        """恢复原始输出"""
        try:
            if self.original_stdout is not None:
                sys.stdout = self.original_stdout
            if self.original_stderr is not None:
                sys.stderr = self.original_stderr
        except (AttributeError, OSError):
            # 在打包环境中可能出现恢复错误，忽略
            pass
    
    def set_log_widget(self, widget):
        """设置日志显示组件"""
        self.log_widget = widget
    
    def set_log_level(self, level):
        """设置日志级别"""
        if level in LogLevel.LEVELS:
            self.current_log_level = level
    
    def should_log(self, level):
        """判断是否应该记录该级别的日志"""
        current_priority = LogLevel.LEVELS.get(self.current_log_level, 20)
        message_priority = LogLevel.LEVELS.get(level, 20)
        return message_priority >= current_priority
    
    def extract_log_level(self, message):
        """从消息中提取日志级别"""
        import re
        # 匹配格式如 [DEBUG] 或 [INFO] 等
        match = re.search(r'\[(\w+)\]', message)
        if match:
            level = match.group(1).upper()
            if level in LogLevel.LEVELS:
                return level
        return None


class ProgressCallback:
    """进度回调类，用于更新GUI进度显示"""
    
    def __init__(self, progress_bar=None, progress_label=None, status_label=None):
        self.progress_bar = progress_bar
        self.progress_label = progress_label
        self.status_label = status_label
        self.total_steps = 0
        self.current_step = 0
    
    def set_total_steps(self, total):
        """设置总步数"""
        self.total_steps = total
        if self.progress_bar:
            self.progress_bar['maximum'] = total
            self.progress_bar['value'] = 0
    
    def update_progress(self, step, message=""):
        """更新进度"""
        self.current_step = step
        
        if self.progress_bar:
            self.progress_bar['value'] = step
            self.progress_bar.update_idletasks()
        
        if self.progress_label:
            if self.total_steps > 0:
                percentage = int((step / self.total_steps) * 100)
                self.progress_label.set(f"进度: {step}/{self.total_steps} ({percentage}%)")
            else:
                self.progress_label.set(f"进度: {step}")
        
        if self.status_label and message:
            self.status_label.set(message)
    
    def complete(self, message="处理完成"):
        """标记完成"""
        if self.progress_bar:
            self.progress_bar['value'] = self.total_steps
        
        if self.progress_label:
            self.progress_label.set("完成")
        
        if self.status_label:
            self.status_label.set(message)
    
    def reset(self):
        """重置进度"""
        self.current_step = 0
        if self.progress_bar:
            self.progress_bar['value'] = 0
        if self.progress_label:
            self.progress_label.set("就绪")


class LogLevel:
    """日志级别常量"""
    DEBUG = "DEBUG"
    INFO = "INFO"
    WARNING = "WARNING"
    ERROR = "ERROR"
    CRITICAL = "CRITICAL"
    
    # 日志级别优先级（数字越大级别越高）
    LEVELS = {
        "DEBUG": 10,
        "INFO": 20,
        "WARNING": 30,
        "ERROR": 40,
        "CRITICAL": 50
    }


class EnhancedLogger:
    """增强的日志记录器"""
    
    def __init__(self, gui_handler=None):
        self.gui_handler = gui_handler
        self.log_history = []
        
        # 设置标准logging
        self.logger = logging.getLogger('ExcelProcessor')
        self.logger.setLevel(logging.DEBUG)
        
        # 创建格式化器
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        
        # 文件处理器
        try:
            file_handler = logging.FileHandler('excel_processor.log', encoding='utf-8')
            file_handler.setFormatter(formatter)
            self.logger.addHandler(file_handler)
        except Exception:
            pass  # 如果无法创建文件日志，继续运行
    
    def log(self, level, message):
        """记录日志"""
        timestamp = datetime.now()
        log_entry = {
            'timestamp': timestamp,
            'level': level,
            'message': message
        }
        self.log_history.append(log_entry)
        
        # 限制历史记录数量
        if len(self.log_history) > 1000:
            self.log_history = self.log_history[-500:]
        
        # 格式化消息
        formatted_message = f"[{level}] {message}"
        
        # 安全地输出到控制台和GUI
        try:
            print(formatted_message)
        except (AttributeError, OSError):
            # 在打包环境中print可能失败，只写入到GUI
            if gui_log_handler.log_widget:
                gui_log_handler.write_to_gui(formatted_message, level)
        
        # 记录到标准日志
        if level == LogLevel.DEBUG:
            self.logger.debug(message)
        elif level == LogLevel.INFO:
            self.logger.info(message)
        elif level == LogLevel.WARNING:
            self.logger.warning(message)
        elif level == LogLevel.ERROR:
            self.logger.error(message)
        elif level == LogLevel.CRITICAL:
            self.logger.critical(message)
    
    def debug(self, message):
        self.log(LogLevel.DEBUG, message)
    
    def info(self, message):
        self.log(LogLevel.INFO, message)
    
    def warning(self, message):
        self.log(LogLevel.WARNING, message)
    
    def error(self, message):
        self.log(LogLevel.ERROR, message)
    
    def critical(self, message):
        self.log(LogLevel.CRITICAL, message)
    
    def get_history(self, level=None, limit=None):
        """获取日志历史"""
        history = self.log_history
        
        if level:
            history = [entry for entry in history if entry['level'] == level]
        
        if limit:
            history = history[-limit:]
        
        return history
    
    def clear_history(self):
        """清空日志历史"""
        self.log_history.clear()


# 全局日志实例
gui_log_handler = GUILogHandler()
enhanced_logger = EnhancedLogger(gui_log_handler) 