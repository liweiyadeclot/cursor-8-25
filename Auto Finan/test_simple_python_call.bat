@echo off
chcp 65001 >nul
echo 简单Python调用测试
echo ==================

echo.
echo 运行简单Python调用测试...
python test_simple_python_call.py --config config.json --folder "C:\Users\FH\Documents\报销单" --file "test_file.pdf"

echo.
echo 测试完成！
pause






