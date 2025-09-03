@echo off
echo ========================================
echo 测试打印按钮Python脚本集成功能
echo ========================================
echo.
echo 正在编译项目...
dotnet build
if %ERRORLEVEL% EQU 0 (
    echo ✅ 编译成功！
    echo.
    echo 🚀 开始运行程序...
    echo 注意：程序会自动查找并点击打印按钮
    echo 然后执行mouse_keyboard_automation.py脚本
    echo.
    echo 请确保mouse_keyboard_automation.py文件存在
    echo.
    dotnet run
) else (
    echo ❌ 编译失败！
    echo 请检查错误信息。
)
pause







