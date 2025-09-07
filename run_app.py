#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel IP地址处理工具 v2.0 启动脚本
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
    print(f"导入错误：{e}")
    print("请确保已安装所需依赖：pip install pandas openpyxl")
    input("按回车键退出...")
except Exception as e:
    print(f"程序运行错误：{e}")
    input("按回车键退出...") 