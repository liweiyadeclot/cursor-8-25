@echo off
chcp 65001 >nul
echo 测试打印确认单按钮集成功能
echo =============================

echo.
echo 正在编译项目...
dotnet build

if %ERRORLEVEL% NEQ 0 (
    echo 编译失败！
    pause
    exit /b 1
)

echo.
echo 编译成功！

echo.
echo 测试Python脚本执行器...
python mouse_keyboard_automation.py --x 1200 --y 800

echo.
echo 测试完成！
pause






