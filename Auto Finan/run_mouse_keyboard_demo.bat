@echo off
chcp 65001 >nul
echo ========================================
echo 鼠标键盘自动化功能演示
echo ========================================
echo.

echo 正在检查Python环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误：未找到Python环境，请先安装Python
    pause
    exit /b 1
)

echo 正在检查依赖包...
python -c "import pyautogui" >nul 2>&1
if errorlevel 1 (
    echo 正在安装pyautogui...
    pip install pyautogui
    if errorlevel 1 (
        echo 错误：安装pyautogui失败
        pause
        exit /b 1
    )
)

echo.
echo 启动鼠标键盘自动化演示...
echo 注意：请确保目标窗口在前台
echo.

python mouse_keyboard_automation.py

echo.
echo 演示完成！
pause
