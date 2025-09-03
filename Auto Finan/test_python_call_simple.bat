@echo off
chcp 65001 >nul
echo 简化Python调用测试
echo ===================

echo.
echo 测试简化Python调用...
python test_python_call_simple.py --config config.json --folder "C:\Users\FH\Documents\报销单" --file "test_file.pdf"

echo.
echo 测试完成！
pause






