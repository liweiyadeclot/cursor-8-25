@echo off
echo 财务自动化系统主程序
echo ========================

echo 正在编译项目...
dotnet build

if %ERRORLEVEL% EQU 0 (
    echo 编译成功！
    echo 正在运行主程序...
    echo.
    echo 程序启动后，您可以选择：
    echo 1. 执行查询程序（科研财务系统）
    echo 2. 执行报销程序（财务报销自动化）
    echo.
    dotnet run --project "Auto Finan.csproj" --configuration Debug
) else (
    echo 编译失败！
    pause
)

echo.
echo 程序执行完成，按任意键退出...
pause




