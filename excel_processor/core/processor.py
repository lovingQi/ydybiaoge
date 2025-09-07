import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from datetime import datetime
import openpyxl
import re
import threading
import time
import sys
import os
from pathlib import Path

# 添加项目根目录到Python路径
current_dir = Path(__file__).parent.parent
sys.path.insert(0, str(current_dir))

from utils.logger import enhanced_logger, ProgressCallback


class ExcelProcessor:
    """Excel文件处理核心类"""
    
    def __init__(self, progress_callback=None):
        self.progress_callback = progress_callback
        self.is_processing = False
        self.should_cancel = False
        self.current_file = None
        self.processing_thread = None
        
    def set_progress_callback(self, callback):
        """设置进度回调"""
        self.progress_callback = callback
    
    def cancel_processing(self):
        """取消处理"""
        self.should_cancel = True
        enhanced_logger.info("用户请求取消处理")
    
    def select_excel_file(self):
        """
        打开文件选择对话框，让用户选择Excel文件
        
        Returns:
            str: 选中的文件路径，如果用户取消则返回None
        """
        # 创建tkinter根窗口（隐藏）
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        
        try:
            # 打开文件选择对话框
            file_path = filedialog.askopenfilename(
                title="选择Excel文件",
                initialdir=os.getcwd(),
                filetypes=[
                    ("Excel文件", "*.xlsx *.xls"),
                    ("Excel 2007-365", "*.xlsx"),
                    ("Excel 97-2003", "*.xls"),
                    ("所有文件", "*.*")
                ]
            )
            
            # 销毁根窗口
            root.destroy()
            
            # 如果用户选择了文件，返回路径；否则返回None
            return file_path if file_path else None
            
        except Exception as e:
            root.destroy()
            enhanced_logger.error(f"文件选择过程中出现错误：{str(e)}")
            return None

    def read_excel_data(self, file_path):
        """
        读取Excel文件第一个工作表的所有数据
        
        Args:
            file_path (str): Excel文件路径
            
        Returns:
            pandas.DataFrame: 包含Excel数据的DataFrame对象，出错时返回None
        """
        try:
            # 检查是否需要取消
            if self.should_cancel:
                enhanced_logger.info("读取文件前检测到取消请求")
                return None
                
            # 检查文件是否存在
            if not os.path.exists(file_path):
                enhanced_logger.error(f"文件不存在：{file_path}")
                return None
            
            enhanced_logger.info(f"正在读取Excel文件：{file_path}")
            
            # 读取Excel文件第一个工作表
            df = pd.read_excel(
                file_path,
                sheet_name=0,  # 读取第一个工作表
                engine='openpyxl'  # 使用openpyxl引擎
            )
            
            # 读取完成后再次检查是否需要取消
            if self.should_cancel:
                enhanced_logger.info("文件读取完成后检测到取消请求")
                return None
            
            enhanced_logger.info(f"文件读取成功，数据形状：{df.shape}")
            return df
            
        except FileNotFoundError:
            enhanced_logger.error(f"找不到指定文件：{file_path}")
            return None
        except pd.errors.EmptyDataError:
            enhanced_logger.error("Excel文件为空或无有效数据")
            return None
        except Exception as e:
            enhanced_logger.error(f"读取Excel文件时出现错误：{str(e)}")
            return None

    def normalize_ip(self, ip_str):
        """
        将IP地址统一为CIDR格式

        Args:
            ip_str (str): 原始IP字符串
            
        Returns:
            str: 标准化后的IP字符串
        """
        if pd.isna(ip_str) or not isinstance(ip_str, str):
            return ip_str
            
        ip_str = ip_str.strip()

        # 情况1: ip/32 格式，如 36.134.205.246/32 - 保持不变
        if re.match(r'^\d+\.\d+\.\d+\.\d+/32$', ip_str):
            return ip_str
            
        # 情况2: 网段/24 格式，如 36.134.41.0/24 - 保持不变
        if re.match(r'^\d+\.\d+\.\d+\.\d+/24$', ip_str):
            return ip_str
            
        # 情况3: 纯IP格式，如 36.134.44.69 - 添加/32
        if re.match(r'^\d+\.\d+\.\d+\.\d+$', ip_str):
            return f"{ip_str}/32"
            
        # 情况4: 网段格式，如 36.134.41.0 - 添加/24
        # 判断是否为网段（第四段为0）
        ip_parts = ip_str.split('.')
        if len(ip_parts) == 4 and ip_parts[3] == '0':
            return f"{ip_str}/24"
            
        # 其他情况保持原样
        return ip_str

    def detect_duplicate_ips(self, df):
        """
        检测DataFrame中第二列IP地址的重复情况
        新逻辑：除最小行号外，所有重复行都需要标记
        
        Args:
            df (pandas.DataFrame): 包含IP数据的DataFrame
            
        Returns:
            dict: 重复信息字典，格式为 {行索引: 标记文本}
        """
        try:
            if self.should_cancel:
                return {}
                
            enhanced_logger.info("开始检测重复IP地址...")
            
            # 提取第二列从第三行开始的IP数据（跳过表头）
            ip_column = df.iloc[2:, 1]  # 从索引2开始，第二列（索引1）
            
            # 提取第五列运营商数据
            isp_column = df.iloc[2:, 4]  # 从索引2开始，第五列（索引4）
            
            # 不重置索引，直接使用原始DataFrame索引
            ip_data = ip_column.dropna()  # 只移除空值，保持原始索引
            
            duplicate_info = {}
            
            # 检查每个IP的重复情况
            for ip in ip_data.unique():
                if self.should_cancel:
                    break
                    
                # 找到所有该IP的DataFrame索引位置
                positions = ip_data[ip_data == ip].index.tolist()
                
                if len(positions) > 1:  # 如果有重复
                    # 转换为Excel行号（DataFrame索引 + 1）
                    excel_rows = [pos + 2 for pos in positions]  # 转换为Excel行号
                    
                    # 按行号排序
                    excel_rows.sort()
                    
                    # 为每个需要标记的行检查运营商
                    for i in range(1, len(excel_rows)):  # 从第二个开始（跳过最小行号）
                        current_row = excel_rows[i]
                        current_row_index = current_row - 2  # 转回DataFrame索引
                        current_isp = isp_column.iloc[current_row_index - 2]  # 当前行的运营商
                        
                        # 检查是否需要标记（与前面行比较运营商）
                        should_mark = False
                        mark_text = ""
                        for j in range(i):
                            prev_row = excel_rows[j]
                            prev_row_index = prev_row - 2  # 转回DataFrame索引
                            prev_isp = isp_column.iloc[prev_row_index - 2]  # 前面行的运营商
                            
                            if (current_isp == prev_isp):
                                should_mark = True
                                # 如果mark_text中间包含"保留"，则先清空
                                if "保留" in mark_text:
                                    mark_text = ""
                                mark_text += f"删除这行，与第{prev_row - 2}行重复，且运营商相同，都是{current_isp};"
                            elif((current_isp == "联通" or current_isp == "电信") and (prev_isp == "电信/联通" or prev_isp == "联通/电信")):
                                should_mark = True
                                 # 如果mark_text中间包含"保留"，则先清空
                                if "保留" in mark_text:
                                    mark_text = ""
                                mark_text += f"删除这行，与第{prev_row - 2}行重复，且运营商{current_isp}被其包含了;"
                            elif((current_isp == "电信/联通" or current_isp == "联通/电信") and (prev_isp == "电信/联通" or prev_isp == "联通/电信")):
                                should_mark = True
                                if "保留" in mark_text:
                                    mark_text = ""
                                mark_text += f"删除这行，与第{prev_row - 2}行重复，且运营商相同，都是{current_isp};"
                            elif((current_isp == "电信/联通" or current_isp == "联通/电信") and (prev_isp == "联通" or prev_isp == "电信")):
                                should_mark = True
                                if not mark_text:
                                    mark_text += f"保留这行，与第{prev_row - 2}行重复，需删去运营商{prev_isp};"
                            else:
                                should_mark = False
                        
                        if should_mark:                     
                            # 保存标记信息
                            duplicate_info[current_row_index] = mark_text
            
            enhanced_logger.info(f"重复IP检测完成，发现 {len(duplicate_info)} 个重复标记")
            return duplicate_info
            
        except Exception as e:
            enhanced_logger.error(f"检测重复IP时出现错误：{str(e)}")
            return {}

    def detect_subnet_containment(self, df):
        """
        检测第二列中的网段是否包含其他IP地址
        
        Args:
            df (pandas.DataFrame): 包含IP数据的DataFrame
            
        Returns:
            dict: 包含关系信息字典，格式为 {行索引: 包含描述}
        """
        try:
            if self.should_cancel:
                return {}
                
            enhanced_logger.info("开始检测网段包含关系...")
            containment_info = {}

            # 提取第二列从第三行开始的IP数据（跳过表头）
            ip_column = df.iloc[2:, 1]  # 从索引2开始，第二列（索引1）
            
            # 提取第五列运营商数据
            isp_column = df.iloc[2:, 4]  # 从索引2开始，第五列（索引4）
            
            # 从第三行开始处理（跳过表头）
            for idx in range(2, len(df)):
                if self.should_cancel:
                    break
                    
                ip_value = ip_column.iloc[idx -2]
                # 检查是否为网段格式 (x.x.x.0/24)
                if re.match(r'^\d+\.\d+\.\d+\.0/24$', ip_value):
                    # 提取网段前缀
                    prefix = ip_value.split('/')[0].rsplit('.', 1)[0] + '.'
                    
                    # 检查后续行是否被该网段包含
                    current_isp = isp_column.iloc[idx -2]

                    for next_idx in range(idx + 1, len(df)):
                        if self.should_cancel:
                            break
                            
                        next_ip = df.iloc[next_idx, 1]
                        should_mark = False
                        mark_text = ""
                        
                        # 跳过空值或非字符串值
                        if pd.isna(next_ip) or not isinstance(next_ip, str):
                            continue
                            
                        next_ip = next_ip.strip()
                        
                        # 检查是否被包含（以网段前缀开头）
                        if next_ip.startswith(prefix):
                            try:
                                next_isp = isp_column.iloc[next_idx -2]
     
                                if (current_isp == next_isp):
                                    should_mark = True
                                    if "保留" in mark_text:
                                        mark_text = ""
                                    mark_text += f"删除这行，与第{idx}行的网段{prefix}重复，且运营商相同，都是{current_isp};"
                                elif((current_isp == "联通" or current_isp == "电信") and (next_isp == "电信/联通" or next_isp == "联通/电信")):
                                    should_mark = True
                                    if not mark_text:
                                        mark_text += f"保留这行，与第{idx}行的网段{prefix}重复，需删去运营商{current_isp};"                                             
                                elif((current_isp == "电信/联通" or current_isp == "联通/电信") and (next_isp == "电信/联通" or next_isp == "联通/电信")):
                                    should_mark = True
                                    if "保留" in mark_text:
                                        mark_text = ""
                                    mark_text += f"删除这行，与第{idx}行的网段{prefix}重复，且运营商相同，都是{current_isp};"
                                elif((current_isp == "电信/联通" or current_isp == "联通/电信") and (next_isp == "联通" or next_isp == "电信")):
                                    should_mark = True
                                    if "保留" in mark_text:
                                        mark_text = ""
                                    mark_text += f"删除这行，与第{idx}行的网段{prefix}重复，且运营商{next_isp}被其包含了;"
                                else:
                                    should_mark = False
                               
                                if should_mark:
                                    containment_info[next_idx] = mark_text
                            except:
                                mark_text = ""
            
            enhanced_logger.info(f"网段包含关系检查完成，发现 {len(containment_info)} 个包含关系")
            return containment_info
            
        except Exception as e:
            enhanced_logger.error(f"检测网段包含关系时出现错误：{str(e)}")
            return {}

    def check_subnet_consistency(self, df):
        """
        检查IP地址的网段是否与第6列的网段一致
        
        Args:
            df (pandas.DataFrame): 包含IP数据的DataFrame
            
        Returns:
            dict: 不一致信息字典，格式为 {行索引: 不一致描述}
        """
        try:
            if self.should_cancel:
                return {}
                
            enhanced_logger.info("开始检查网段一致性...")
            inconsistency_info = {}
            
            # 从第三行开始处理（跳过表头）
            for idx in range(2, len(df)):
                if self.should_cancel:
                    break
                    
                # 获取第二列的IP地址
                ip_value = df.iloc[idx, 1]
                # 获取第六列的网段值
                subnet_value = df.iloc[idx, 5]
                
                # 跳过空值
                if pd.isna(ip_value) or pd.isna(subnet_value) or not isinstance(ip_value, str) or not isinstance(subnet_value, str):
                    continue
                    
                ip_value = ip_value.strip()
                subnet_value = subnet_value.strip()
                
                # 计算或提取IP的网段
                calculated_subnet = ""
                
                # 情况1: IP/32格式，需要计算网段
                if re.match(r'^\d+\.\d+\.\d+\.\d+/32$', ip_value):
                    ip_parts = ip_value.split('/')[0].split('.')
                    calculated_subnet = f"{ip_parts[0]}.{ip_parts[1]}.{ip_parts[2]}.0/24"
                
                # 情况2: 网段/24格式，直接使用
                elif re.match(r'^\d+\.\d+\.\d+\.\d+/24$', ip_value):
                    calculated_subnet = ip_value
                
                # 其他格式，尝试处理
                else:
                    # 尝试提取IP部分
                    ip_match = re.match(r'^(\d+\.\d+\.\d+\.\d+)', ip_value)
                    if ip_match:
                        ip_parts = ip_match.group(1).split('.')
                        calculated_subnet = f"{ip_parts[0]}.{ip_parts[1]}.{ip_parts[2]}.0/24"
                
                # 检查计算出的网段与第6列是否一致
                if calculated_subnet and calculated_subnet != subnet_value:
                    inconsistency_info[idx] = f"IP网段不一致: {calculated_subnet} ≠ {subnet_value}"
                    
            enhanced_logger.info(f"网段一致性检查完成，发现 {len(inconsistency_info)} 个不一致")
            return inconsistency_info
            
        except Exception as e:
            enhanced_logger.error(f"检查网段一致性时出现错误：{str(e)}")
            return {}

    def preprocess_excel_data(self, df):
        """
        对Excel数据进行预处理，主要是标准化IP地址格式
        
        Args:
            df (pandas.DataFrame): 原始Excel数据
            
        Returns:
            tuple: (预处理后的DataFrame, IP变更字典)
        """
        try:
            if self.should_cancel:
                return df, {}
                
            enhanced_logger.info("开始预处理Excel数据...")
            
            # 复制DataFrame以避免修改原始数据
            processed_df = df.copy()
            
            # 假设IP地址在第二列（索引为1）
            ip_changes = {}
            if processed_df.shape[1] > 1:
                # 从第三行开始处理（跳过表头）
                ip_column = processed_df.iloc[2:, 1]
                
                # 应用normalize_ip函数标准化IP地址并记录处理情况
                for idx in ip_column.index:
                    if self.should_cancel:
                        break
                        
                    original_ip = ip_column.loc[idx]
                    normalized_ip = self.normalize_ip(original_ip)
                    if original_ip != normalized_ip:
                        ip_changes[idx] = f"{original_ip} → {normalized_ip}"
                    processed_df.iloc[idx, 1] = normalized_ip
                
                enhanced_logger.info(f"IP地址格式标准化完成，共处理 {len(ip_changes)} 个IP地址")
            else:
                enhanced_logger.warning("数据列数不足，无法处理IP地址")
                
            return processed_df, ip_changes
            
        except Exception as e:
            enhanced_logger.error(f"预处理数据时出现错误：{str(e)}")
            return df, {}  # 出错时返回原始数据

    def process_excel_file_async(self, file_path, callback=None):
        """
        异步处理Excel文件
        
        Args:
            file_path (str): Excel文件路径
            callback (function): 完成回调函数
        """
        def process_thread():
            try:
                self.is_processing = True
                self.should_cancel = False
                self.current_file = file_path
                
                if self.progress_callback:
                    self.progress_callback.set_total_steps(7)
                    self.progress_callback.update_progress(0, "开始处理...")
                
                # 步骤1: 读取文件
                enhanced_logger.info(f"开始处理文件: {file_path}")
                df = self.read_excel_data(file_path)
                if df is None:
                    if self.should_cancel:
                        enhanced_logger.info("处理已取消")
                        if callback:
                            callback(False, {'error': '用户取消了操作'})
                    return
                if self.should_cancel:
                    enhanced_logger.info("处理已取消")
                    if callback:
                        callback(False, {'error': '用户取消了操作'})
                    return
                
                if self.progress_callback:
                    self.progress_callback.update_progress(1, "文件读取完成")
                
                # 步骤2: 预处理数据
                df, ip_changes = self.preprocess_excel_data(df)
                if self.should_cancel:
                    enhanced_logger.info("预处理阶段检测到取消请求")
                    if callback:
                        callback(False, {'error': '用户取消了操作'})
                    return
                
                if self.progress_callback:
                    self.progress_callback.update_progress(2, "数据预处理完成")
                
                # 步骤3: 检查网段一致性
                inconsistency_info = self.check_subnet_consistency(df)
                if self.should_cancel:
                    enhanced_logger.info("网段一致性检查阶段检测到取消请求")
                    if callback:
                        callback(False, {'error': '用户取消了操作'})
                    return
                
                if self.progress_callback:
                    self.progress_callback.update_progress(3, "网段一致性检查完成")
                
                # 步骤4: 检测重复IP
                duplicate_info = self.detect_duplicate_ips(df)
                if self.should_cancel:
                    enhanced_logger.info("重复IP检测阶段检测到取消请求")
                    if callback:
                        callback(False, {'error': '用户取消了操作'})
                    return
                
                if self.progress_callback:
                    self.progress_callback.update_progress(4, "重复IP检测完成")
                
                # 步骤5: 检测网段包含关系
                subnet_containment_info = self.detect_subnet_containment(df)
                if self.should_cancel:
                    enhanced_logger.info("网段包含关系检测阶段检测到取消请求")
                    if callback:
                        callback(False, {'error': '用户取消了操作'})
                    return
                
                if self.progress_callback:
                    self.progress_callback.update_progress(5, "网段包含关系检测完成")
                
                # 步骤6: 保存处理结果
                success, new_file_path = self.save_processed_excel_v2(
                    ip_changes, inconsistency_info, duplicate_info, 
                    subnet_containment_info, file_path
                )
                
                if self.should_cancel:
                    enhanced_logger.info("文件保存阶段检测到取消请求")
                    if callback:
                        callback(False, {'error': '用户取消了操作'})
                    return
                
                if self.progress_callback:
                    self.progress_callback.update_progress(6, "文件保存完成")
                
                # 步骤7: 完成
                if success:
                    enhanced_logger.info(f"处理完成！新文件已保存为：{new_file_path}")
                    if self.progress_callback:
                        self.progress_callback.complete("处理完成")
                else:
                    enhanced_logger.error("文件保存失败")
                
                # 调用回调函数
                if callback:
                    callback(success, {
                        'ip_changes': len(ip_changes),
                        'inconsistencies': len(inconsistency_info),
                        'duplicates': len(duplicate_info),
                        'containments': len(subnet_containment_info),
                        'output_file': new_file_path if success else None
                    })
                
            except Exception as e:
                enhanced_logger.error(f"处理过程中出现错误：{str(e)}")
                if callback:
                    callback(False, {'error': str(e)})
            finally:
                # 确保状态正确重置
                was_cancelled = self.should_cancel
                self.is_processing = False
                self.should_cancel = False
                
                # 如果是取消操作，确保回调被调用
                if was_cancelled and callback:
                    enhanced_logger.info("处理线程结束时检测到取消状态")
                    callback(False, {'error': '用户取消了操作'})
        
        # 启动处理线程
        self.processing_thread = threading.Thread(target=process_thread)
        self.processing_thread.daemon = True
        self.processing_thread.start()

    def save_processed_excel_v2(self, ip_changes, inconsistency_info, duplicate_info, subnet_containment_info, original_file_path):
        """
        使用openpyxl精确修改Excel文件，保持原有格式
        新增第19列作为"处理方案"列进行标记
        
        Args:
            ip_changes (dict): IP变更字典
            inconsistency_info (dict): 不一致信息字典
            duplicate_info (dict): 重复信息字典
            subnet_containment_info (dict): 网段包含信息字典
            original_file_path (str): 原文件路径
            
        Returns:
            tuple: (成功状态, 新文件路径)
        """
        try:
            if self.should_cancel:
                return False, None
                
            # 生成新文件名
            file_dir = os.path.dirname(original_file_path)
            file_name = os.path.basename(original_file_path)
            name_without_ext, ext = os.path.splitext(file_name)
            new_file_name = f"{name_without_ext}_checked{ext}"
            new_file_path = os.path.join(file_dir, new_file_name)

            new_file_name2 = f"{name_without_ext}_processed{ext}"
            new_file_path2 = os.path.join(file_dir, new_file_name2)

            new_file_name3 = f"{name_without_ext}_result{ext}"
            new_file_path3 = os.path.join(file_dir, new_file_name3)
            
            # 使用openpyxl打开原文件
            enhanced_logger.info(f"正在打开原文件：{original_file_path}")
            wb = openpyxl.load_workbook(original_file_path)
            ws = wb.active
            
            enhanced_logger.info(f"原文件信息：工作表名='{ws.title}', 行数={ws.max_row}, 列数={ws.max_column}")
            
            # 处理IP地址变更，修改第2列的IP地址
            enhanced_logger.info(f"开始处理IP地址变更，共有 {len(ip_changes)} 个需要修改")
            
            # 遍历IP变更记录
            for row_index, change_info in ip_changes.items():
                if self.should_cancel:
                    return False, None
                    
                # pandas行索引转换为Excel行号（+1，因为Excel是1-based）
                excel_row = row_index + 2
                
                # 从变更信息中提取标准化后的IP地址
                normalized_ip = change_info.split(" → ")[1]
                
                # 写入标准化后的IP地址到第2列
                ws.cell(row=excel_row, column=2, value=normalized_ip)
            
            if ip_changes:
                enhanced_logger.info(f"IP地址变更处理完成，共修改 {len(ip_changes)} 个IP地址")
            else:
                enhanced_logger.info("没有需要修改的IP地址")

            # 处理网段一致性检查结果，将不一致信息添加到第19列
            enhanced_logger.info(f"开始处理网段一致性检查结果，共有 {len(inconsistency_info)} 个不一致")
            ws.cell(row=1, column=19, value="网段合法检查")
            ws.cell(row=2, column=19, value=f"共有 {len(inconsistency_info)} 个不一致")
            
            # 遍历网段一致性检查结果
            for row_index, inconsistency_text in inconsistency_info.items():
                if self.should_cancel:
                    return False, None
                    
                # pandas行索引转换为Excel行号（+1，因为Excel是1-based）
                excel_row = row_index + 2
                
                # 获取当前单元格的值
                current_value = ws.cell(row=excel_row, column=19).value
                
                # 如果单元格已有内容，则追加新内容；否则直接写入
                if current_value:
                    ws.cell(row=excel_row, column=19, value=f"{current_value}; {inconsistency_text}")
                else:
                    ws.cell(row=excel_row, column=19, value=inconsistency_text)
                    
            if inconsistency_info:
                enhanced_logger.info(f"网段一致性检查结果处理完成，共标记 {len(inconsistency_info)} 个不一致")
            else:
                enhanced_logger.info("没有发现网段不一致的情况")

            # 确定目标列为第20列（IP重复检查）
            target_column = 20  # 第20列（T列）
            
            # 设置表头：在第1行第20列写入"ip重复检查"
            ws.cell(row=1, column=target_column, value="ip重复检查")
            ws.cell(row=2, column=target_column, value=f"共 {len(duplicate_info)} 处IP重复")
            
            # 修改需要标记的单元格
            for row_index, mark_text in duplicate_info.items():
                if self.should_cancel:
                    return False, None
                    
                # pandas行索引转换为Excel行号（+1，因为Excel是1-based）
                excel_row = row_index + 2
                
                # 写入标记文本到第20列
                ws.cell(row=excel_row, column=target_column, value=mark_text)

            # 处理网段包含检查
            ws.cell(row=1, column=21, value="网段包含检查")
            ws.cell(row=2, column=21, value=f"共有 {len(subnet_containment_info)} 个包含关系")
            for row_index, mark_text in subnet_containment_info.items():
                if self.should_cancel:
                    return False, None
                    
                excel_row = row_index + 2
                ws.cell(row=excel_row, column=21, value=mark_text)

            if subnet_containment_info:
                enhanced_logger.info(f"网段包含检查结果处理完成，共标记 {len(subnet_containment_info)} 个包含关系")
            else:
                enhanced_logger.info("没有发现网段包含的情况")

            # 保存检查结果文件
            enhanced_logger.info(f"正在保存检查结果文件：{new_file_path}")
            wb.save(new_file_path)

            # 开始具体处理19列（网段不一致）
            enhanced_logger.info("开始处理网段不一致情况...")
            
            # 从第3行开始检查第19列
            for row in range(3, ws.max_row + 1):
                if self.should_cancel:
                    return False, None
                    
                # 获取第19列的内容
                inconsistency_text = ws.cell(row=row, column=19).value
                
                # 检查单元格是否有内容且包含"IP网段不一致"
                if inconsistency_text and isinstance(inconsistency_text, str) and "IP网段不一致" in inconsistency_text:
                    # 使用正则表达式提取第一个网段（通常是正确的网段）
                    match = re.search(r'(\d+\.\d+\.\d+\.\d+/\d+)\s*≠', inconsistency_text)
                    
                    if match:
                        correct_subnet = match.group(1)
                        # 将提取的网段填入第6列（F列）
                        ws.cell(row=row, column=6, value=correct_subnet)
            
            enhanced_logger.info("网段不一致处理完成")

            # 处理第20列的重复标记
            enhanced_logger.info("开始处理第20列的重复标记...")
            
            # 从第3行开始检查第20列
            for row in range(3, ws.max_row + 1):
                if self.should_cancel:
                    return False, None
                    
                # 获取第20列的内容
                duplicate_text = ws.cell(row=row, column=20).value
                
                # 检查单元格是否有内容
                if duplicate_text and isinstance(duplicate_text, str):
                    # 处理"保留这行"的情况
                    if "保留这行" in duplicate_text:
                        # 提取需要删除的运营商
                        match = re.search(r'需删去运营商([\u4e00-\u9fa5]{2});', duplicate_text)
                        if match:
                            isp_to_remove = match.group(1)
                            
                            # 获取当前运营商值（第5列）
                            current_isp = ws.cell(row=row, column=5).value
                            if current_isp and isinstance(current_isp, str):
                                # 删除指定运营商和斜杠
                                new_isp = current_isp.replace(f"/{isp_to_remove}", "").replace(f"{isp_to_remove}/", "")
                                # 更新运营商值
                                ws.cell(row=row, column=5).value = new_isp
                    
                    # 处理"删除这行"的情况
                    elif "删除这行" in duplicate_text:
                        # 清空该行前18列的内容
                        for col in range(1, 19):
                            ws.cell(row=row, column=col).value = None
            
            enhanced_logger.info("第20列重复标记处理完成")

            # 处理第21列的网段包含标记
            enhanced_logger.info("开始处理第21列的网段包含标记...")
            
            for row in range(3, ws.max_row + 1):
                if self.should_cancel:
                    return False, None
                    
                duplicate_text = ws.cell(row=row, column=21).value
                
                if duplicate_text and isinstance(duplicate_text, str):
                    # 处理"保留这行"的情况
                    if "保留这行" in duplicate_text:
                        match = re.search(r'需删去运营商([\u4e00-\u9fa5]{2});', duplicate_text)
                        if match:
                            isp_to_remove = match.group(1)
                            
                            current_isp = ws.cell(row=row, column=5).value
                            if current_isp and isinstance(current_isp, str):
                                new_isp = current_isp.replace(f"/{isp_to_remove}", "").replace(f"{isp_to_remove}/", "")
                                ws.cell(row=row, column=5).value = new_isp
                    
                    # 处理"删除这行"的情况
                    elif "删除这行" in duplicate_text:
                        for col in range(1, 19):
                            ws.cell(row=row, column=col).value = None
            
            enhanced_logger.info("第21列网段包含标记处理完成")

            # 保存处理后的文件
            enhanced_logger.info(f"正在保存处理后文件：{new_file_path2}")
            wb.save(new_file_path2)

            # 清除19、20、21、22列的内容 - 优化版本
            enhanced_logger.info("开始清除19、20、21、22列的内容...")
            
            # 删除整列而不是逐个清除单元格，这样更高效
            columns_to_delete = []
            for col in range(22, 18, -1):  # 从右到左删除列，避免索引变化
                if col <= ws.max_column:
                    columns_to_delete.append(col)
            
            for col in columns_to_delete:
                if self.should_cancel:
                    return False, None
                ws.delete_cols(col)
                
            enhanced_logger.info("19、20、21、22列已删除")
            
            # 删除第1列内容为空的行 - 优化版本
            enhanced_logger.info("开始删除第1列内容为空的行...")
            
            # 批量删除空行，提高效率
            rows_to_delete = []
            
            # 先收集所有需要删除的行号
            for row in range(3, ws.max_row + 1):  # 从第3行开始检查，保留表头
                if self.should_cancel:
                    return False, None
                    
                cell_value = ws.cell(row=row, column=1).value
                if cell_value is None or (isinstance(cell_value, str) and cell_value.strip() == ""):
                    rows_to_delete.append(row)
            
            # 从后往前删除，避免索引变化
            total_to_delete = len(rows_to_delete)
            for i, row in enumerate(reversed(rows_to_delete)):
                if self.should_cancel:
                    return False, None
                ws.delete_rows(row)
                
                # 每删除50行输出一次进度
                if total_to_delete > 50 and (i + 1) % 50 == 0:
                    remaining = total_to_delete - (i + 1)
                    enhanced_logger.info(f"删除进度: 已删除 {i + 1}/{total_to_delete} 行，还剩 {remaining} 行...")
            
            enhanced_logger.info(f"共删除了{len(rows_to_delete)}行（第1列为空的行）")

            # 保存最终结果文件
            enhanced_logger.info(f"正在保存最终结果文件：{new_file_path3}")
            wb.save(new_file_path3)

            wb.close()
            
            enhanced_logger.info(f"文件处理完成，共生成3个文件：")
            enhanced_logger.info(f"1. 检查结果: {new_file_path}")
            enhanced_logger.info(f"2. 处理过程: {new_file_path2}")
            enhanced_logger.info(f"3. 最终结果: {new_file_path3}")
            
            return True, new_file_path3
            
        except Exception as e:
            enhanced_logger.error(f"使用openpyxl保存文件时出现错误：{str(e)}")
            return False, None 