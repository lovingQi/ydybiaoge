import sys
import io
from datetime import datetime
import logging


class GUILogHandler:
    """GUI日志处理器，将print输出重定向到GUI界面"""
    
    def __init__(self, log_widget=None):
        self.log_widget = log_widget
        self.original_stdout = sys.stdout
        self.original_stderr = sys.stderr
        self.log_buffer = io.StringIO()
        
    def write(self, text):
        """重定向的写入方法"""
        # 写入到原始输出（保持控制台输出）
        self.original_stdout.write(text)
        self.original_stdout.flush()
        
        # 如果有GUI组件，也写入到GUI
        if self.log_widget and text.strip():
            self.write_to_gui(text.strip())
    
    def flush(self):
        """刷新缓冲区"""
        self.original_stdout.flush()
    
    def write_to_gui(self, message):
        """将消息写入GUI日志组件"""
        if self.log_widget:
            try:
                timestamp = datetime.now().strftime("%H:%M:%S")
                formatted_message = f"[{timestamp}] {message}\n"
                
                # 启用文本框编辑
                self.log_widget.config(state='normal')
                
                # 插入新消息
                self.log_widget.insert('end', formatted_message)
                
                # 自动滚动到底部
                self.log_widget.see('end')
                
                # 禁用文本框编辑
                self.log_widget.config(state='disabled')
                
                # 更新界面
                self.log_widget.update_idletasks()
                
            except Exception as e:
                # 如果GUI写入失败，至少保证控制台输出
                self.original_stdout.write(f"GUI日志写入失败: {e}\n")
    
    def redirect_output(self):
        """开始重定向输出"""
        sys.stdout = self
        sys.stderr = self
    
    def restore_output(self):
        """恢复原始输出"""
        sys.stdout = self.original_stdout
        sys.stderr = self.original_stderr
    
    def set_log_widget(self, widget):
        """设置日志显示组件"""
        self.log_widget = widget


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
        
        # 输出到控制台和GUI
        print(formatted_message)
        
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