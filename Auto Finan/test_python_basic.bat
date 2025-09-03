@echo off
chcp 65001 >nul
echo 基本Python测试
echo ===============

echo.
echo 测试基本Python脚本...
python test_python_basic.py --config config.json --folder "C:\Users\FH\Documents\报销单" --file "test_file.pdf"

echo.
echo 测试完成！
pause






