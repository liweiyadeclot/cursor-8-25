@echo off
chcp 65001 >nul
echo Python问题诊断工具
echo ===================

echo.
echo 开始诊断Python相关问题...
echo.

echo 1. 检查Python环境...
python --version
if %ERRORLEVEL% NEQ 0 (
    echo ✗ Python未安装或不在PATH中
    echo 请安装Python并确保在PATH环境变量中
    goto :end
)
echo ✓ Python环境正常
echo.

echo 2. 检查必要文件...
if exist "test_mouse_keyboard.py" (
    echo ✓ test_mouse_keyboard.py 存在
) else (
    echo ✗ test_mouse_keyboard.py 不存在
)

if exist "config.json" (
    echo ✓ config.json 存在
) else (
    echo ✗ config.json 不存在
)

if exist "test_simple_python_call.py" (
    echo ✓ test_simple_python_call.py 存在
) else (
    echo ✗ test_simple_python_call.py 不存在
)
echo.

echo 3. 测试简单Python调用...
python test_simple_python_call.py --config config.json --folder "C:\Users\FH\Documents\报销单" --file "test_file.pdf"
if %ERRORLEVEL% NEQ 0 (
    echo ✗ 简单Python调用失败
    goto :end
)
echo ✓ 简单Python调用成功
echo.

echo 4. 检查Python包...
python -c "import pyautogui; print('✓ pyautogui 已安装')" 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo ✗ pyautogui 未安装
    echo 请运行: pip install pyautogui
    goto :end
)
echo ✓ pyautogui 已安装
echo.

echo 5. 测试完整脚本...
python test_mouse_keyboard.py --help
if %ERRORLEVEL% NEQ 0 (
    echo ✗ 完整脚本测试失败
    goto :end
)
echo ✓ 完整脚本测试成功
echo.

echo ========================================
echo 诊断完成！所有检查都通过了。
echo 如果仍然有问题，请检查：
echo 1. 配置文件中的坐标是否正确
echo 2. 是否有权限访问相关目录
echo 3. 防火墙是否阻止了Python脚本
echo ========================================

:end
echo.
echo 诊断完成！
pause






