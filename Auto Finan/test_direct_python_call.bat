@echo off
chcp 65001 >nul
echo 测试直接Python调用
echo ==================

echo.
echo 方法1: 使用批处理直接调用...
python test_mouse_keyboard.py --config config.json --folder "C:\Users\FH\Documents\报销单" --file "test_file.pdf"

echo.
echo 方法2: 使用Python脚本测试调用...
python test_direct_python_call.py

echo.
echo 测试完成！
pause






