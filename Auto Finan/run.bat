@echo off
echo === 财务报销自动化系统 ===
echo.

echo 正在检查.NET环境...
dotnet --version
if %errorlevel% neq 0 (
    echo 错误: 未找到.NET环境，请先安装.NET 8.0
    pause
    exit /b 1
)

echo.
echo 正在还原NuGet包...
dotnet restore
if %errorlevel% neq 0 (
    echo 错误: NuGet包还原失败
    pause
    exit /b 1
)

echo.
echo 正在构建项目...
dotnet build
if %errorlevel% neq 0 (
    echo 错误: 项目构建失败
    pause
    exit /b 1
)

echo.
echo 正在安装Playwright浏览器...
pwsh bin\Debug\net8.0\playwright.ps1 install
if %errorlevel% neq 0 (
    echo 警告: Playwright浏览器安装失败，请手动安装
    echo 运行命令: pwsh bin\Debug\net8.0\playwright.ps1 install
)

echo.
echo 正在运行程序...
dotnet run

echo.
echo 程序执行完成
pause

