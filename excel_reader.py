import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from datetime import datetime
import openpyxl
import re  


def select_excel_file():
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
        messagebox.showerror("错误", f"文件选择过程中出现错误：{str(e)}")
        return None


def read_excel_data(file_path):
    """
    读取Excel文件第一个工作表的所有数据
    
    Args:
        file_path (str): Excel文件路径
        
    Returns:
        pandas.DataFrame: 包含Excel数据的DataFrame对象，出错时返回None
    """
    try:
        # 检查文件是否存在
        if not os.path.exists(file_path):
            messagebox.showerror("错误", f"文件不存在：{file_path}")
            return None
        
        # 读取Excel文件第一个工作表
        df = pd.read_excel(
            file_path,
            sheet_name=0,  # 读取第一个工作表
            engine='openpyxl'  # 使用openpyxl引擎
        )
        
        return df
        
    except FileNotFoundError:
        messagebox.showerror("错误", f"找不到指定文件：{file_path}")
        return None
    except pd.errors.EmptyDataError:
        messagebox.showerror("错误", "Excel文件为空或无有效数据")
        return None
    except Exception as e:
        messagebox.showerror("错误", f"读取Excel文件时出现错误：{str(e)}")
        return None


def load_excel_with_gui():
    """
    整合文件选择和数据读取功能的主函数
    
    Returns:
        pandas.DataFrame: 包含Excel数据的DataFrame对象，出错或取消时返回None
    """
    # 选择文件
    file_path = select_excel_file()
    
    if file_path is None:
        print("用户取消了文件选择")
        return None
    
    print(f"选择的文件：{file_path}")
    
    # 读取数据
    df = read_excel_data(file_path)
    
    if df is not None:
        print(f"成功读取Excel文件，数据形状：{df.shape}")
        print("数据预览：")
        print(df.head())
        return df
    else:
        print("读取Excel文件失败")
        return None
 # 处理IP地址格式，统一为CIDR格式
def normalize_ip(ip_str):
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

def detect_duplicate_ips(df):
    """
    检测DataFrame中第二列IP地址的重复情况
    新逻辑：除最小行号外，所有重复行都需要标记
    
    Args:
        df (pandas.DataFrame): 包含IP数据的DataFrame
        
    Returns:
        dict: 重复信息字典，格式为 {行索引: 标记文本}
    """
    try:
        # 提取第二列从第三行开始的IP数据（跳过表头）
        ip_column = df.iloc[2:, 1]  # 从索引2开始，第二列（索引1）
        
        # 提取第五列运营商数据
        isp_column = df.iloc[2:, 4]  # 从索引2开始，第五列（索引4）
        
        
        # 不重置索引，直接使用原始DataFrame索引
        ip_data = ip_column.dropna()  # 只移除空值，保持原始索引
        
        duplicate_info = {}
        
        # 检查每个IP的重复情况
        for ip in ip_data.unique():
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
        
        return duplicate_info
        
    except Exception as e:
        print(f"检测重复IP时出现错误：{str(e)}")
        return {}

def detect_subnet_containment(df):
    """
    检测第二列中的网段是否包含其他IP地址
    
    Args:
        df (pandas.DataFrame): 包含IP数据的DataFrame
        
    Returns:
        dict: 包含关系信息字典，格式为 {行索引: 包含描述}
    """
    try:
        containment_info = {}

        # 提取第二列从第三行开始的IP数据（跳过表头）
        ip_column = df.iloc[2:, 1]  # 从索引2开始，第二列（索引1）
        
        # 提取第五列运营商数据
        isp_column = df.iloc[2:, 4]  # 从索引2开始，第五列（索引4）

        # ip_data = ip_column.dropna()  # 只移除空值，保持原始索引

        
        # 从第三行开始处理（跳过表头）
        for idx in range(2, len(df)):
            # 获取第二列的值
            # ip_value = df.iloc[idx, 1]
            # # 获取第5列的运营商
            # isp_column = df.iloc[idx, 4]
            # 跳过空值或非字符串值
            # if pd.isna(ip_value) or not isinstance(ip_value, str):
            #     continue
                
            # ip_value = ip_value.strip()
            ip_value = ip_column.iloc[idx -2]
            # 检查是否为网段格式 (x.x.x.0/24)
            if re.match(r'^\d+\.\d+\.\d+\.0/24$', ip_value):
                # 提取网段前缀
                prefix = ip_value.split('/')[0].rsplit('.', 1)[0] + '.'
                # 验证: 如果ip_value是36.133.125.0/24，prefix将是36.133.125.
                
                # 检查后续行是否被该网段包含
                current_isp = isp_column.iloc[idx -2]

                print(f"检查网段包含关系: 当前行={idx+2}, 网段={ip_value}, 前缀={prefix}")
                for next_idx in range(idx + 1, len(df)):
                    next_ip = df.iloc[next_idx, 1]
                     # 检查是否需要标记（与前面行比较运营商）
                    should_mark = False
                    mark_text = ""
                    # 跳过空值或非字符串值
                    if pd.isna(next_ip) or not isinstance(next_ip, str):
                        continue
                        
                    next_ip = next_ip.strip()
                    
                    # 检查是否被包含（以网段前缀开头）
                    if next_ip.startswith(prefix):
                        # 获取运营商信息（如果有的话）
                        # 打印调试信息
                        print(f"检查网段包含关系: 当前行={idx+2}, 网段={ip_value}, 前缀={prefix}")
                        print(f"比较行={next_idx+2}, IP={next_ip}")
                        try:
                            next_isp = isp_column.iloc[next_idx -2]
 
                            if (current_isp == next_isp):
                                should_mark = True
                                # 如果mark_text中间包含"保留"，则先清空
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
                                # 如果mark_text中间包含"保留"，则先清空
                                if "保留" in mark_text:
                                    mark_text = ""
                                mark_text += f"删除这行，与第{idx}行的网段{prefix}重复，且运营商{next_isp}被其包含了;"
                            else:
                                should_mark = False
                           
                            if should_mark:
                                containment_info[next_idx] = mark_text
                        except:
                            mark_text = ""
                            
                        # # 保存包含关系信息
                        # containment_info[next_idx] = f"IP {next_ip} 被第{idx+2}行的网段 {ip_value} 包含{isp_info}"
                        # print(f"发现包含关系：行{next_idx+2}的IP {next_ip} 被行{idx+2}的网段 {ip_value} 包含")
        
        print(f"网段包含关系检查完成，发现 {len(containment_info)} 个包含关系")
        return containment_info
        
    except Exception as e:
        print(f"检测网段包含关系时出现错误：{str(e)}")
        return {}




def add_duplicate_marks(df, duplicate_info):
    """
    在DataFrame副本的最后一列添加重复标记
    
    Args:
        df (pandas.DataFrame): 原始DataFrame
        duplicate_info (dict): 重复信息字典
        
    Returns:
        pandas.DataFrame: 添加标记后的DataFrame副本
    """
    try:
        # 创建DataFrame副本
        df_copy = df.copy()
        
        # 在最后一列添加重复标记
        for row_index, mark_text in duplicate_info.items():
            df_copy.iloc[row_index, -1] = mark_text
        
        return df_copy
        
    except Exception as e:
        print(f"添加重复标记时出现错误：{str(e)}")
        return df


def save_processed_excel(df, original_file_path):
    """
    保存处理后的Excel文件
    
    Args:
        df (pandas.DataFrame): 处理后的DataFrame
        original_file_path (str): 原文件路径
        
    Returns:
        tuple: (成功状态, 新文件路径)
    """
    try:
        # 生成新文件名
        file_dir = os.path.dirname(original_file_path)
        file_name = os.path.basename(original_file_path)
        name_without_ext, ext = os.path.splitext(file_name)
        new_file_name = f"{name_without_ext}_checked{ext}"
        new_file_path = os.path.join(file_dir, new_file_name)
        
        # 保存文件
        df.to_excel(new_file_path, index=False, engine='openpyxl')
        
        return True, new_file_path
        
    except Exception as e:
        print(f"保存文件时出现错误：{str(e)}")
        return False, None


def show_duplicate_statistics(duplicate_info):
    """
    显示重复检查的统计信息
    
    Args:
        duplicate_info (dict): 重复信息字典
    """
    print("\n" + "="*50)
    print("IP重复检查统计报告")
    print("="*50)
    
    if not duplicate_info:
        print("✅ 未发现重复的IP地址")
    else:
        print(f"⚠️  发现 {len(duplicate_info)} 个重复标记位置")
        print("\n详细信息：")
        
        for row_index, mark_text in sorted(duplicate_info.items()):
            excel_row = row_index + 1
            print(f"  第{excel_row}行: {mark_text}")
    
    print("="*50)


def get_excel_column_letter(col_index):
    """
    将列索引转换为Excel列字母
    
    Args:
        col_index (int): 列索引（0-based）
        
    Returns:
        str: Excel列字母（如：0→A, 25→Z, 26→AA）
    """
    result = ""
    while col_index >= 0:
        result = chr(65 + col_index % 26) + result
        col_index = col_index // 26 - 1
    return result
def preprocess_excel_data(df):
    """
    对Excel数据进行预处理，主要是标准化IP地址格式
    
    Args:
        df (pandas.DataFrame): 原始Excel数据
        
    Returns:
        tuple: (预处理后的DataFrame, IP变更字典)
    """
    try:
        # 复制DataFrame以避免修改原始数据
        processed_df = df.copy()
        
        # 假设IP地址在第二列（索引为1）
        ip_changes = {}
        if processed_df.shape[1] > 1:
            # 从第三行开始处理（跳过表头）
            ip_column = processed_df.iloc[2:, 1]
            
            # 应用normalize_ip函数标准化IP地址并记录处理情况
            for idx in ip_column.index:
                original_ip = ip_column.loc[idx]
                normalized_ip = normalize_ip(original_ip)
                if original_ip != normalized_ip:
                    ip_changes[idx] = f"{original_ip} → {normalized_ip}"
                processed_df.iloc[idx, 1] = normalized_ip
            
            print(f"IP地址格式标准化完成，共处理 {len(ip_changes)} 个IP地址")
            
            print("IP地址格式标准化完成")
        else:
            print("警告：数据列数不足，无法处理IP地址")
            
        return processed_df, ip_changes
        
    except Exception as e:
        print(f"预处理数据时出现错误：{str(e)}")
        return df, ip_changes  # 出错时返回原始数据和已处理的变更
def check_subnet_consistency(df):
    """
    检查IP地址的网段是否与第6列的网段一致
    
    Args:
        df (pandas.DataFrame): 包含IP数据的DataFrame
        
    Returns:
        dict: 不一致信息字典，格式为 {行索引: 不一致描述}
    """
    try:
        inconsistency_info = {}
        
        # 从第三行开始处理（跳过表头）
        for idx in range(2, len(df)):
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
                
        print(f"网段一致性检查完成，发现 {len(inconsistency_info)} 个不一致")
        return inconsistency_info
        
    except Exception as e:
        print(f"检查网段一致性时出现错误：{str(e)}")
        return {}


def save_processed_excel_v2(ip_changes, inconsistency_info, duplicate_info, subnet_containment_info, original_file_path):
    """
    使用openpyxl精确修改Excel文件，保持原有格式
    新增第19列作为"处理方案"列进行标记
    
    Args:
        duplicate_info (dict): 重复信息字典
        original_file_path (str): 原文件路径
        
    Returns:
        tuple: (成功状态, 新文件路径)
    """
    try:
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
        print(f"正在打开原文件：{original_file_path}")
        wb = openpyxl.load_workbook(original_file_path)
        ws = wb.active
        
        print(f"原文件信息：工作表名='{ws.title}', 行数={ws.max_row}, 列数={ws.max_column}")
        
        # 处理IP地址变更，修改第2列的IP地址
        print(f"开始处理IP地址变更，共有 {len(ip_changes)} 个需要修改")


 #-----------------------------------------------------------------------------
        
        # 遍历IP变更记录
        for row_index, change_info in ip_changes.items():
            # pandas行索引转换为Excel行号（+1，因为Excel是1-based）
            excel_row = row_index + 2
            
            # 从变更信息中提取标准化后的IP地址
            normalized_ip = change_info.split(" → ")[1]
            
            # 写入标准化后的IP地址到第2列
            ws.cell(row=excel_row, column=2, value=normalized_ip)
            print(f"修改第{excel_row}行第2列IP地址：{change_info}")
        
        if ip_changes:
            print(f"IP地址变更处理完成，共修改 {len(ip_changes)} 个IP地址")
        else:
            print("没有需要修改的IP地址")

#-----------------------------------------------------------------------------

        # 处理网段一致性检查结果，将不一致信息添加到第19列
        print(f"开始处理网段一致性检查结果，共有 {len(inconsistency_info)} 个不一致")
        ws.cell(row=1, column=19, value="网段合法检查")
        ws.cell(row=2, column=19, value=f"共有 {len(inconsistency_info)} 个不一致")
        # 遍历网段一致性检查结果
        for row_index, inconsistency_text in inconsistency_info.items():
            # pandas行索引转换为Excel行号（+1，因为Excel是1-based）
            excel_row = row_index + 2
            
            # 获取当前单元格的值
            current_value = ws.cell(row=excel_row, column=19).value
            
            # 如果单元格已有内容，则追加新内容；否则直接写入
            if current_value:
                ws.cell(row=excel_row, column=19, value=f"{current_value}; {inconsistency_text}")
            else:
                ws.cell(row=excel_row, column=19, value=inconsistency_text)
                
            print(f"标记第{excel_row}行第19列网段不一致：{inconsistency_text}")
            
        if inconsistency_info:
            print(f"网段一致性检查结果处理完成，共标记 {len(inconsistency_info)} 个不一致")
        else:
            print("没有发现网段不一致的情况")

#-----------------------------------------------------------------------------

        
        # 确定目标列为第19列（新增的处理方案列）
        target_column = 20  # 第20列（S列）
        
        # 设置表头：在第1行第19列写入"处理方案"
        ws.cell(row=1, column=target_column, value="ip重复检查")
        print(f"设置表头：第1行第{target_column}列 = 'ip重复检查'")

        ws.cell(row=2, column=target_column, value=f"共 {len(duplicate_info)} 处IP重复")
        # 修改需要标记的单元格
        for row_index, mark_text in duplicate_info.items():
            # pandas行索引转换为Excel行号（+1，因为Excel是1-based）
            excel_row = row_index + 2
            
            # 写入标记文本到第19列
            ws.cell(row=excel_row, column=target_column, value=mark_text)
            print(f"标记第{excel_row}行第{target_column}列：{mark_text}")
        
#-----------------------------------------------------------------------------

        ws.cell(row=1, column=21, value="网段包含检查")
        ws.cell(row=2, column=21, value=f"共有 {len(subnet_containment_info)} 个包含关系")
        for row_index, mark_text in subnet_containment_info.items():
            excel_row = row_index + 2
            ws.cell(row=excel_row, column=21, value=mark_text)
            print(f"标记第{excel_row}行第21列：{mark_text}")

        if subnet_containment_info:
            print(f"网段包含检查结果处理完成，共标记 {len(subnet_containment_info)} 个包含关系")
        else:
            print("没有发现网段包含的情况")

#-----------------------------------------------------------------------------


        # 保存文件
        print(f"正在保存新文件：{new_file_path}")
        wb.save(new_file_path)


#-----------------------------------------------------------------------------开始做具体处理19列
        # 处理第19列的网段不一致情况，提取正确网段填入第6列
        print("开始处理网段不一致情况...")
        
        # 从第3行开始检查第19列
        for row in range(3, ws.max_row + 1):
            # 获取第19列的内容
            inconsistency_text = ws.cell(row=row, column=19).value
            
            # 检查单元格是否有内容且包含"IP网段不一致"
            if inconsistency_text and isinstance(inconsistency_text, str) and "IP网段不一致" in inconsistency_text:
                print(f"发现第{row}行网段不一致：{inconsistency_text}")
                
                # 使用正则表达式提取第一个网段（通常是正确的网段）
                # 匹配格式如：36.134.42.0/24 ≠ 36.134.43.0/24
                match = re.search(r'(\d+\.\d+\.\d+\.\d+/\d+)\s*≠', inconsistency_text)
                
                if match:
                    correct_subnet = match.group(1)
                    print(f"提取到正确网段：{correct_subnet}")
                    
                    # 将提取的网段填入第6列（F列）
                    ws.cell(row=row, column=6, value=correct_subnet)
                    print(f"已将正确网段 {correct_subnet} 填入第{row}行第6列")
                else:
                    print(f"无法从文本中提取网段：{inconsistency_text}")
        
        print("网段不一致处理完成")
        # # 清空第19列的内容
        # print("开始清空第19列内容...")
        # for row in range(1, ws.max_row + 1):
        #     # 清空第19列的单元格内容
        #     ws.cell(row=row, column=19).value = None
        
        # print("第19列内容已全部清空")
#-----------------------------------------------------------------------------开始做具体处理20列
        # 处理第20列的重复标记
        print("开始处理第20列的重复标记...")
        
        # 从第3行开始检查第20列
        for row in range(3, ws.max_row + 1):
            # 获取第20列的内容
            duplicate_text = ws.cell(row=row, column=20).value
            
            # 检查单元格是否有内容
            if duplicate_text and isinstance(duplicate_text, str):
                print(f"处理第{row}行重复标记：{duplicate_text}")
                
                # 处理"保留这行"的情况
                if "保留这行" in duplicate_text:
                    # 提取需要删除的运营商
                    match = re.search(r'需删去运营商([\u4e00-\u9fa5]{2});', duplicate_text)
                    if match:
                        isp_to_remove = match.group(1)
                        print(f"需要从运营商中删除：{isp_to_remove}")
                        
                        # 获取当前运营商值（第5列）
                        current_isp = ws.cell(row=row, column=5).value
                        if current_isp and isinstance(current_isp, str):
                            # 删除指定运营商和斜杠
                            new_isp = current_isp.replace(f"/{isp_to_remove}", "").replace(f"{isp_to_remove}/", "")
                            # 更新运营商值
                            ws.cell(row=row, column=5).value = new_isp
                            print(f"已将运营商从 {current_isp} 更新为 {new_isp}")
                
                # 处理"删除这行"的情况
                elif "删除这行" in duplicate_text:
                    print(f"清空第{row}行前18列内容")
                    # 清空该行前18列的内容
                    for col in range(1, 19):
                        ws.cell(row=row, column=col).value = None
        
        print("第20列重复标记处理完成")

#-----------------------------------------------------------------------------开始做具体处理21列

        # 处理第20列的重复标记
        print("开始处理第21列的重复标记...")
        
        # 从第3行开始检查第20列
        for row in range(3, ws.max_row + 1):
            # 获取第20列的内容
            duplicate_text = ws.cell(row=row, column=21).value
            
            # 检查单元格是否有内容
            if duplicate_text and isinstance(duplicate_text, str):
                print(f"处理第{row}行重复标记：{duplicate_text}")
                
                # 处理"保留这行"的情况
                if "保留这行" in duplicate_text:
                    # 提取需要删除的运营商
                    match = re.search(r'需删去运营商([\u4e00-\u9fa5]{2});', duplicate_text)
                    if match:
                        isp_to_remove = match.group(1)
                        print(f"需要从运营商中删除：{isp_to_remove}")
                        
                        # 获取当前运营商值（第5列）
                        current_isp = ws.cell(row=row, column=5).value
                        if current_isp and isinstance(current_isp, str):
                            # 删除指定运营商和斜杠
                            new_isp = current_isp.replace(f"/{isp_to_remove}", "").replace(f"{isp_to_remove}/", "")
                            # 更新运营商值
                            ws.cell(row=row, column=5).value = new_isp
                            print(f"已将运营商从 {current_isp} 更新为 {new_isp}")
                
                # 处理"删除这行"的情况
                elif "删除这行" in duplicate_text:
                    print(f"清空第{row}行前18列内容")
                    # 清空该行前18列的内容
                    for col in range(1, 19):
                        ws.cell(row=row, column=col).value = None
        
        print("第21列重复标记处理完成")


        print(f"正在保存新文件：{new_file_path2}")
        wb.save(new_file_path2)

#-----------------------------------------------------------------------------

        # 清除19、20、21、22列的内容
        print("开始清除19、20、21、22列的内容...")
        for row in range(1, ws.max_row + 1):
            for col in range(19, 23):  # 19、20、21、22列
                if col <= ws.max_column:  # 确保列存在
                    ws.cell(row=row, column=col).value = None
        print("19、20、21、22列内容已清除")
        
        # 删除第1列内容为空的行
        print("开始删除第1列内容为空的行...")
        rows_to_delete = []
        
        # 先标记需要删除的行
        for row in range(ws.max_row, 0, -1):  # 从底部向上遍历，避免删除时索引变化
            cell_value = ws.cell(row=row, column=1).value
            if cell_value is None or (isinstance(cell_value, str) and cell_value.strip() == ""):
                rows_to_delete.append(row)
        
        # 删除标记的行
        for row in rows_to_delete:
            ws.delete_rows(row)
            print(f"已删除第{row}行（第1列为空）")
        
        print(f"共删除了{len(rows_to_delete)}行（第1列为空的行）")




        print(f"正在保存新文件：{new_file_path3}")
        wb.save(new_file_path3)

        wb.close()
        
        print(f"文件保存成功：{new_file_path}")
        print(f"新文件列数应为：{ws.max_column} (原{ws.max_column-1}列 + 新增1列)")
        return True, new_file_path
        
    except Exception as e:
        print(f"使用openpyxl保存文件时出现错误：{str(e)}")
        return False, None


def check_and_mark_duplicates_v2():
    """
    使用openpyxl精确保存的重复检查主函数
    
    Returns:
        bool: 处理成功返回True，失败返回False
    """
    print("启动IP重复检查功能（openpyxl版本）...")
    
    # 选择并读取Excel文件
    file_path = select_excel_file()
    if file_path is None:
        print("用户取消了文件选择")
        return False
    
    print(f"读取文件：{file_path}")
    df = read_excel_data(file_path)
    if df is None:
        print("文件读取失败")
        return False
    
    print(f"文件读取成功，数据形状：{df.shape}")
    
    # 预处理数据
    print("正在预处理数据...")
    df, ip_changes = preprocess_excel_data(df)

    inconsistency_info = check_subnet_consistency(df)
    
    # 检测重复IP
    print("正在检测重复IP...")
    duplicate_info = detect_duplicate_ips(df)


    #检测网段包含关系
    print("正在检测网段包含关系...")
    subnet_containment_info = detect_subnet_containment(df)
    
    # 显示统计信息
    show_duplicate_statistics(duplicate_info)
    
    if duplicate_info:
        # 使用openpyxl精确保存
        print("正在使用openpyxl精确保存文件...")
        success, new_file_path = save_processed_excel_v2(ip_changes, inconsistency_info, duplicate_info, subnet_containment_info, file_path)
        
        if success:
            print(f"✅ 处理完成！新文件已保存为：{new_file_path}")
            
            # 显示成功消息框
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo("处理完成", 
                              f"IP重复检查完成！\n\n"
                              f"发现 {len(duplicate_info)} 个重复标记\n"
                              f"文件已保存为：\n{os.path.basename(new_file_path)}")
            root.destroy()
            
            return True
        else:
            print("❌ 文件保存失败")
            return False
    else:
        print("✅ 未发现重复IP，无需处理")
        return True


# 使用示例和测试代码
if __name__ == "__main__":
    print("Excel文件读取程序启动...")
    print("1. 基本读取功能测试")
    check_and_mark_duplicates_v2()

    # # 调用基本读取功能
    # data = load_excel_with_gui()
    
    # if data is not None:
    #     print("\n程序执行成功！")
    #     print(f"数据类型：{type(data)}")
    #     print(f"数据维度：{data.shape}")
    #     print(f"列名：{list(data.columns)}")
        
    #     print("\n2. IP重复检查功能测试（openpyxl精确保存版本）")
    #     # 调用新的重复检查功能
    #     check_and_mark_duplicates_v2()
    # else:
    #     print("\n程序执行结束，未获取到数据") 