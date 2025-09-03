@echo off
echo ========================================
echo 测试编译修复
echo ========================================
echo.
echo 正在编译项目...
dotnet build
if %ERRORLEVEL% EQU 0 (
    echo ✅ 编译成功！
    echo.
    echo 🎉 编译问题已修复！
    echo 现在可以正常运行程序了。
    echo.
    echo 要运行程序，请执行：
    echo dotnet run
) else (
    echo ❌ 编译仍然失败！
    echo 请检查错误信息。
)
pause







