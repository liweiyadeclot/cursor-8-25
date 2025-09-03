@echo off
chcp 65001 >nul
echo 测试Python脚本 test_mouse_keyboard.py
echo =====================================

echo.
echo 检查Python环境...
python --version

if %ERRORLEVEL% NEQ 0 (
    echo Python未安装或不在PATH中！
    pause
    exit /b 1
)

echo.
echo 检查配置文件...
if not exist "config.json" (
    echo 错误：config.json文件不存在！
    pause
    exit /b 1
)

echo.
echo 检查Python脚本...
if not exist "test_mouse_keyboard.py" (
    echo 错误：test_mouse_keyboard.py文件不存在！
    pause
    exit /b 1
)

echo.
echo 开始测试Python脚本...
echo 将在5秒后开始执行，请将鼠标移动到屏幕角落可随时中断...
echo.

python test_mouse_keyboard.py --config config.json --folder "C:\Users\FH\Documents\报销单" --file "test_file.pdf"

echo.
echo 测试完成！
pause






