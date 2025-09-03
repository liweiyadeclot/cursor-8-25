@echo off
chcp 65001 >nul
echo 测试Python脚本执行器
echo ====================

echo.
echo 正在编译项目...
dotnet build

if %ERRORLEVEL% NEQ 0 (
    echo 编译失败！
    pause
    exit /b 1
)

echo.
echo 编译成功！正在运行测试...
echo.

dotnet run --project "Auto Finan.csproj"

echo.
echo 测试完成！
pause
