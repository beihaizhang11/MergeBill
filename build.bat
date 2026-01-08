@echo off
chcp 65001 >nul
echo ========================================
echo Excel账单合并工具 - 打包脚本
echo ========================================
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

echo [步骤2/3] 开始打包程序...
echo 这可能需要几分钟时间，请耐心等待...
echo.

pyinstaller --noconfirm ^
    --onefile ^
    --windowed ^
    --name "Excel账单合并工具" ^
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
echo 可执行文件位置: dist\Excel账单合并工具.exe
echo 文件大小: 
for %%F in ("dist\Excel账单合并工具.exe") do echo %%~zF 字节 (约 %%~zF/1048576 MB)
echo.
echo 注意事项：
echo 1. 首次运行会在同目录生成config.json配置文件
echo 2. 配置文件可以手动备份和迁移
echo 3. 建议将exe文件放在独立文件夹中使用
echo.
pause

