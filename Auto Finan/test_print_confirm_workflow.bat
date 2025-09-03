@echo off
chcp 65001 >nul
echo 测试打印确认单按钮完整工作流程
echo =================================

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
echo 测试Python脚本执行器（打印对话框自动化）...
python mouse_keyboard_automation.py

echo.
echo 测试完成！
echo.
echo 说明：
echo 1. 当在Excel中设置"打印确认单"列的值为"$点击"时
echo 2. 程序会通过Playwright点击网页上的打印确认单按钮
echo 3. 点击成功后，会自动调用Python脚本处理后续的鼠标键盘操作
echo 4. Python脚本会等待并处理可能弹出的打印对话框
echo.
pause






