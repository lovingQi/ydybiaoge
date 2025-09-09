#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel IP地址处理工具 v2.1
主程序入口

功能：
1. IP地址格式标准化
2. 重复IP检测和处理
3. 网段包含关系检测
4. 网段一致性检查
"""

import sys
import os
import tkinter as tk
from tkinter import messagebox
import threading
from pathlib import Path

# 添加项目根目录到Python路径
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from gui.main_window import MainWindow
from core.processor import ExcelProcessor
from utils.logger import gui_log_handler, enhanced_logger, ProgressCallback


class ExcelProcessorApp:
    """Excel处理器应用程序主类"""
    
    def __init__(self):
        self.main_window = None
        self.processor = None
        self.progress_callback = None
        
    def initialize(self):
        """初始化应用程序"""
        try:
            # 创建主窗口
            self.main_window = MainWindow()
            
            # 设置日志处理器
            gui_log_handler.set_log_widget(self.main_window.log_text)
            gui_log_handler.configure_text_colors(self.main_window.log_text)
            
            # 从配置文件读取日志级别
            config_log_level = self.main_window.config.get("logging", {}).get("default_level", "INFO")
            gui_log_handler.set_log_level(config_log_level)
            self.main_window.log_level_var.set(config_log_level)
            
            gui_log_handler.redirect_output()
            
            # 创建进度回调
            self.progress_callback = ProgressCallback(
                progress_bar=self.main_window.progress_bar,
                progress_label=self.main_window.progress_var,
                status_label=self.main_window.status_var
            )
            
            # 创建处理器
            self.processor = ExcelProcessor(self.progress_callback)
            
            # 绑定事件处理
            self.bind_events()
            
            enhanced_logger.info("Excel IP地址处理工具 v2.1 初始化完成")
            return True
            
        except Exception as e:
            error_msg = f"应用程序初始化失败：{str(e)}"
            try:
                print(error_msg)
            except:
                pass  # 在打包环境中print可能失败
            try:
                messagebox.showerror("初始化错误", error_msg)
            except:
                pass  # GUI可能还未初始化
            return False
    
    def bind_events(self):
        """绑定事件处理"""
        # 重新绑定主窗口的处理按钮
        self.main_window.process_btn.config(command=self.start_processing)
        self.main_window.cancel_btn.config(command=self.cancel_processing)
        
        # 绑定文件选择
        original_select_file = self.main_window.select_file
        def enhanced_select_file():
            file_path = self.processor.select_excel_file()
            if file_path:
                self.main_window.file_var.set(os.path.basename(file_path))
                self.main_window.current_file = file_path
                self.main_window.process_btn.config(state=tk.NORMAL)
                enhanced_logger.info(f"已选择文件: {file_path}")
                self.main_window.status_var.set(f"已选择文件: {os.path.basename(file_path)}")
        
        self.main_window.select_file = enhanced_select_file
        self.main_window.select_btn.config(command=enhanced_select_file)
        
        # 更新菜单中的文件选择命令
        self.main_window.update_file_menu_command(enhanced_select_file)
    
    def start_processing(self):
        """开始处理文件"""
        if not hasattr(self.main_window, 'current_file') or not self.main_window.current_file:
            messagebox.showwarning("警告", "请先选择要处理的Excel文件")
            return
        
        # 更新UI状态
        self.main_window.process_btn.config(state=tk.DISABLED)
        self.main_window.cancel_btn.config(state=tk.NORMAL)
        self.main_window.select_btn.config(state=tk.DISABLED)
        
        # 重置进度
        self.progress_callback.reset()
        
        # 清空统计信息
        self.update_stats_display("")
        
        enhanced_logger.info("="*50)
        enhanced_logger.info("开始处理Excel文件")
        enhanced_logger.info("="*50)
        
        # 异步处理文件
        self.processor.process_excel_file_async(
            self.main_window.current_file,
            self.on_processing_complete
        )
    
    def cancel_processing(self):
        """取消处理"""
        if self.processor and self.processor.is_processing:
            self.processor.cancel_processing()
            enhanced_logger.info("正在取消处理...")
            
            # 更新UI状态
            self.main_window.cancel_btn.config(state=tk.DISABLED)
            self.main_window.status_var.set("正在取消...")
            
            # 重置进度条
            self.progress_callback.reset()
            self.main_window.progress_var.set("已取消")
    
    def on_processing_complete(self, success, results):
        """处理完成回调"""
        # 恢复UI状态
        self.main_window.process_btn.config(state=tk.NORMAL)
        self.main_window.cancel_btn.config(state=tk.DISABLED)
        self.main_window.select_btn.config(state=tk.NORMAL)
        
        if success:
            enhanced_logger.info("="*50)
            enhanced_logger.info("处理完成！")
            enhanced_logger.info("="*50)
            
            # 更新统计信息
            stats_text = self.format_results(results)
            self.update_stats_display(stats_text)
            
            # 添加到历史记录
            self.add_to_history(success, results)
            
            # 显示完成消息
            messagebox.showinfo("处理完成", 
                              f"文件处理完成！\n\n"
                              f"IP地址变更: {results.get('ip_changes', 0)} 个\n"
                              f"网段不一致: {results.get('inconsistencies', 0)} 个\n"
                              f"重复IP: {results.get('duplicates', 0)} 个\n"
                              f"网段包含: {results.get('containments', 0)} 个\n\n"
                              f"结果文件: {os.path.basename(results.get('output_file', ''))}")
        else:
            error_msg = results.get('error', '未知错误')
            if '用户取消了操作' in error_msg:
                enhanced_logger.info("用户已取消处理")
                self.update_stats_display("处理已取消")
                self.main_window.status_var.set("已取消")
                # 不显示错误对话框，因为这是用户主动取消的
            else:
                enhanced_logger.error("处理失败")
                self.update_stats_display(f"处理失败: {error_msg}")
                messagebox.showerror("处理失败", f"文件处理失败：{error_msg}")
            
            self.add_to_history(success, results)
    
    def format_results(self, results):
        """格式化处理结果"""
        stats = []
        stats.append("处理结果统计:")
        stats.append("-" * 20)
        stats.append(f"IP地址变更: {results.get('ip_changes', 0)} 个")
        stats.append(f"网段不一致: {results.get('inconsistencies', 0)} 个")
        stats.append(f"重复IP标记: {results.get('duplicates', 0)} 个")
        stats.append(f"网段包含关系: {results.get('containments', 0)} 个")
        stats.append("")
        
        if results.get('output_file'):
            stats.append("输出文件:")
            stats.append(os.path.basename(results['output_file']))
        
        return "\n".join(stats)
    
    def update_stats_display(self, text):
        """更新统计信息显示"""
        self.main_window.stats_text.config(state=tk.NORMAL)
        self.main_window.stats_text.delete(1.0, tk.END)
        self.main_window.stats_text.insert(1.0, text)
        self.main_window.stats_text.config(state=tk.DISABLED)
    
    def add_to_history(self, success, results):
        """添加到历史记录"""
        import datetime
        
        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        file_name = os.path.basename(self.main_window.current_file) if hasattr(self.main_window, 'current_file') else "未知文件"
        status = "成功" if success else "失败"
        
        if success:
            result_summary = f"IP变更:{results.get('ip_changes', 0)} 不一致网段:{results.get('inconsistencies', 0)} 重复IP:{results.get('duplicates', 0)} 网段包含:{results.get('containments', 0)}"
        else:
            result_summary = results.get('error', '处理失败')
        
        # 插入到历史记录树形视图
        self.main_window.history_tree.insert('', 0, values=(file_name, current_time, status, result_summary))
        
        # 限制历史记录数量
        items = self.main_window.history_tree.get_children()
        if len(items) > 50:  # 保留最近50条记录
            for item in items[50:]:
                self.main_window.history_tree.delete(item)
    
    def run(self):
        """运行应用程序"""
        if self.initialize():
            try:
                self.main_window.run()
            except KeyboardInterrupt:
                enhanced_logger.info("用户中断程序")
            except Exception as e:
                enhanced_logger.error(f"程序运行时出现错误：{str(e)}")
            finally:
                self.cleanup()
        else:
            sys.exit(1)
    
    def cleanup(self):
        """清理资源"""
        try:
            # 恢复标准输出
            gui_log_handler.restore_output()
            
            # 如果有正在进行的处理，取消它
            if self.processor and self.processor.is_processing:
                self.processor.cancel_processing()
                
            enhanced_logger.info("应用程序已退出")
            
        except Exception as e:
            print(f"清理资源时出现错误：{str(e)}")


def main():
    """主函数"""
    try:
        # 设置异常处理
        def handle_exception(exc_type, exc_value, exc_traceback):
            if issubclass(exc_type, KeyboardInterrupt):
                sys.__excepthook__(exc_type, exc_value, exc_traceback)
                return
            
            error_msg = f"未处理的异常：{exc_type.__name__}: {exc_value}"
            try:
                print(error_msg)
            except:
                pass  # 在打包环境中print可能失败
            
            # 如果GUI可用，显示错误对话框
            try:
                root = tk.Tk()
                root.withdraw()
                messagebox.showerror("程序错误", error_msg)
                root.destroy()
            except:
                pass
        
        sys.excepthook = handle_exception
        
        # 创建并运行应用程序
        app = ExcelProcessorApp()
        app.run()
        
    except Exception as e:
        error_msg = f"程序启动失败：{str(e)}"
        try:
            print(error_msg)
        except:
            pass  # 在打包环境中print可能失败
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("启动错误", error_msg)
            root.destroy()
        except:
            pass
        sys.exit(1)


if __name__ == "__main__":
    main() 