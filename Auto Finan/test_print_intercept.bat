@echo off
echo ========================================
echo 测试报销确认单打印拦截功能
echo ========================================
echo.
echo 正在编译项目...
dotnet build
if %ERRORLEVEL% EQU 0 (
    echo ✅ 编译成功！
    echo.
    echo 🚀 开始运行程序...
    echo 注意：程序会自动查找并点击报销确认单按钮
    echo 然后尝试拦截打印内容并保存为HTML文件
    echo.
    dotnet run
) else (
    echo ❌ 编译失败！
    echo 请检查错误信息。
)
pause







