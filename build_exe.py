#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel IP地址处理工具 v2.1 打包脚本
使用PyInstaller将程序打包为exe文件
"""

import os
import sys
import subprocess
from pathlib import Path

def build_exe():
    """构建exe文件"""
    
    print("开始打包Excel IP地址处理工具...")
    
    # 确保在正确的目录
    project_root = Path(__file__).parent
    os.chdir(project_root)
    
    # PyInstaller命令参数
    cmd = [
        'pyinstaller',
        '--onefile',                    # 打包成单个exe文件
        '--windowed',                   # 不显示控制台窗口
        '--name=ExcelIPProcessor',      # exe文件名
        '--icon=icon.ico',              # 图标文件（如果有的话）
        '--add-data=excel_processor;excel_processor',  # 包含整个excel_processor目录
        '--hidden-import=pandas',       # 确保pandas被包含
        '--hidden-import=openpyxl',     # 确保openpyxl被包含
        '--hidden-import=tkinter',      # 确保tkinter被包含
        '--hidden-import=excel_processor.gui.main_window',  # 显式包含GUI模块
        '--hidden-import=excel_processor.core.processor',   # 显式包含核心处理模块
        '--hidden-import=excel_processor.utils.logger',     # 显式包含日志模块
        '--hidden-import=excel_processor.gui.styles',       # 显式包含样式模块
        '--clean',                      # 清理临时文件
        'run_app.py'                    # 主程序入口
    ]
    
    # 如果没有图标文件，移除图标参数
    if not os.path.exists('icon.ico'):
        cmd = [arg for arg in cmd if not arg.startswith('--icon')]
    
    try:
        print("执行打包命令...")
        print(" ".join(cmd))
        
        # 执行打包命令
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        
        print("打包成功！")
        print(f"输出: {result.stdout}")
        
        # 检查生成的exe文件
        exe_path = project_root / "dist" / "ExcelIPProcessor.exe"
        if exe_path.exists():
            print(f"✅ exe文件已生成: {exe_path}")
            print(f"文件大小: {exe_path.stat().st_size / (1024*1024):.1f} MB")
        else:
            print("❌ exe文件未找到")
            
    except subprocess.CalledProcessError as e:
        print(f"打包失败: {e}")
        print(f"错误输出: {e.stderr}")
        return False
    except FileNotFoundError:
        print("❌ PyInstaller未安装，请先运行: pip install pyinstaller")
        return False
    
    return True

def create_spec_file():
    """创建PyInstaller规格文件，用于更精确的控制"""
    
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['run_app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('excel_processor', 'excel_processor'),
    ],
    hiddenimports=[
        'pandas',
        'openpyxl',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'excel_processor.gui.main_window',
        'excel_processor.core.processor',
        'excel_processor.utils.logger',
        'excel_processor.gui.styles',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ExcelIPProcessor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico' if os.path.exists('icon.ico') else None,
)
'''
    
    with open('ExcelIPProcessor.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)
    
    print("✅ 已创建 ExcelIPProcessor.spec 文件")

def build_from_spec():
    """使用spec文件打包"""
    try:
        cmd = ['pyinstaller', '--clean', 'ExcelIPProcessor.spec']
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("使用spec文件打包成功！")
        return True
    except subprocess.CalledProcessError as e:
        print(f"使用spec文件打包失败: {e}")
        print(f"错误输出: {e.stderr}")
        return False

if __name__ == "__main__":
    print("Excel IP地址处理工具 v2.0 打包工具")
    print("=" * 50)
    
    # # 检查依赖
    # try:
    #     import pyinstaller
    #     print("✅ PyInstaller 已安装")
    # except ImportError:
    #     print("❌ PyInstaller 未安装")
    #     print("请运行: pip install pyinstaller")
    #     sys.exit(1)
    
    # 选择打包方式
    print("\n选择打包方式:")
    print("1. 直接打包 (推荐)")
    print("2. 创建spec文件后打包 (高级)")
    
    choice = input("请选择 (1/2): ").strip()
    
    if choice == "2":
        create_spec_file()
        if input("是否立即使用spec文件打包? (y/n): ").lower() == 'y':
            build_from_spec()
    else:
        build_exe()
    
    print("\n打包完成！")
    print("生成的文件在 dist/ 目录中") 