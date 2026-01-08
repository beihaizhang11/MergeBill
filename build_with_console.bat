@echo off
chcp 65001 >nul
echo ========================================
echo Excel账单合并工具 - 打包脚本（调试版本）
echo ========================================
echo.
echo 此版本会显示控制台窗口，用于查看调试信息
echo.

REM 检查是否安装了PyInstaller
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo [提示] 未检测到PyInstaller，正在安装...
    pip install pyinstaller
    echo.
)

echo [步骤1/3] 清理旧的打包文件...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "*.spec" del /q *.spec
echo 清理完成！
echo.

echo [步骤2/3] 开始打包程序（带控制台）...
echo.

pyinstaller --noconfirm ^
    --onefile ^
    --console ^
    --name "Excel账单合并工具_调试版" ^
    --icon=NONE ^
    --add-data "config.json;." ^
    --hidden-import openpyxl ^
    --hidden-import openpyxl.cell ^
    --hidden-import openpyxl.styles ^
    main.py

if errorlevel 1 (
    echo.
    echo [错误] 打包失败！请检查错误信息。
    pause
    exit /b 1
)

echo.
echo [步骤3/3] 清理临时文件...
if exist "build" rmdir /s /q build
if exist "*.spec" del /q *.spec

echo.
echo ========================================
echo 打包完成！
echo ========================================
echo.
echo 可执行文件位置: dist\Excel账单合并工具_调试版.exe
echo.
echo 此版本会显示控制台窗口，便于查看错误信息和调试
echo.
pause

