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
pyinstaller --onefile --windowed --name=ExcelIPProcessor --add-data="excel_processor/resources;excel_processor/resources" --hidden-import=pandas --hidden-import=openpyxl --hidden-import=tkinter --clean run_app.py

echo 打包完成！
echo exe文件位置: dist\ExcelIPProcessor.exe

pause 