@echo off
chcp 65001 >nul
echo 测试Enter键功能
echo ===============

echo.
echo 运行Enter键测试...
echo 注意：请确保当前焦点在文件路径输入框中
echo.
python test_enter_key.py

echo.
echo 测试完成！
pause






