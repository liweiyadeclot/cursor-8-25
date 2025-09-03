@echo off
echo ========================================
echo 测试配置文件功能
echo ========================================
echo.
echo 正在编译项目...
dotnet build
if %ERRORLEVEL% EQU 0 (
    echo ✅ 编译成功！
    echo.
    echo 🚀 开始运行程序...
    echo 注意：程序会从config.json读取配置
    echo 如果找不到配置文件，会使用默认配置
    echo.
    dotnet run
) else (
    echo ❌ 编译失败！
    echo 请检查错误信息。
)
pause







