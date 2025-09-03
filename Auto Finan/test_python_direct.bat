@echo off
chcp 65001 >nul
echo 直接测试Python脚本
echo ===================

echo.
echo 测试Python脚本是否能正常运行...
echo.

echo 1. 检查Python版本...
python --version
if %ERRORLEVEL% NEQ 0 (
    echo ✗ Python未安装或不在PATH中
    goto :end
)
echo ✓ Python环境正常
echo.

echo 2. 检查文件是否存在...
if exist "test_mouse_keyboard.py" (
    echo ✓ test_mouse_keyboard.py 存在
) else (
    echo ✗ test_mouse_keyboard.py 不存在
    goto :end
)

if exist "config.json" (
    echo ✓ config.json 存在
) else (
    echo ✗ config.json 不存在
    goto :end
)
echo.

echo 3. 测试Python脚本帮助信息...
python test_mouse_keyboard.py --help
if %ERRORLEVEL% NEQ 0 (
    echo ✗ Python脚本测试失败
    goto :end
)
echo ✓ Python脚本可以正常运行
echo.

echo 4. 直接执行Python脚本（模拟调用）...
echo 执行命令: python test_mouse_keyboard.py --config config.json --folder "C:\Users\FH\Documents\报销单" --file "test_file.pdf"
echo.
python test_mouse_keyboard.py --config config.json --folder "C:\Users\FH\Documents\报销单" --file "test_file.pdf"
echo.
echo 退出码: %ERRORLEVEL%

if %ERRORLEVEL% EQU 0 (
    echo ✓ Python脚本执行成功
) else (
    echo ✗ Python脚本执行失败
)

:end
echo.
echo 测试完成！
pause






