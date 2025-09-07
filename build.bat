@echo off
echo Excel IP地址处理工具 v2.0 打包脚本
echo =====================================

echo 检查PyInstaller...
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo 正在安装PyInstaller...
    pip install pyinstaller
)

echo 开始打包...
pyinstaller --onefile --windowed --name=ExcelIPProcessor --add-data="excel_processor;excel_processor" --hidden-import=pandas --hidden-import=openpyxl --hidden-import=tkinter --hidden-import=excel_processor.gui.main_window --hidden-import=excel_processor.core.processor --hidden-import=excel_processor.utils.logger --hidden-import=excel_processor.gui.styles --clean run_app.py

echo 打包完成！
echo exe文件位置: dist\ExcelIPProcessor.exe

pause 