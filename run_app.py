#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel IP地址处理工具 v2.1 启动脚本
"""

import sys
import os

# 添加项目路径
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

# 导入并运行主程序
try:
    from excel_processor.main import main
    main()
except ImportError as e:
    try:
        print(f"导入错误：{e}")
        print("请确保已安装所需依赖：pip install pandas openpyxl")
        input("按回车键退出...")
    except:
        # 在打包环境中可能无法使用print和input
        import tkinter as tk
        from tkinter import messagebox
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("导入错误", f"导入错误：{e}\n请确保已安装所需依赖")
            root.destroy()
        except:
            pass
except Exception as e:
    try:
        print(f"程序运行错误：{e}")
        input("按回车键退出...")
    except:
        # 在打包环境中可能无法使用print和input
        import tkinter as tk
        from tkinter import messagebox
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("程序错误", f"程序运行错误：{e}")
            root.destroy()
        except:
            pass 