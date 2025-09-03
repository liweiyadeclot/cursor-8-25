@echo off
chcp 65001 >nul
echo 最终修复测试
echo ==============

echo.
echo 1. 测试Unicode编码修复...
python test_unicode_fix.py
echo.

echo 2. 测试Python脚本基本功能...
python test_mouse_keyboard.py --help
echo.

echo 3. 测试配置文件读取...
python -c "import json; config=json.load(open('config.json')); print('SaveFolderPath:', config.get('SaveFolderPath', 'NOT FOUND'))"
echo.

echo 4. 测试完整调用（不执行实际操作）...
python test_mouse_keyboard.py --config config.json --folder "C:\Users\FH\Documents\报销单" --file "test_file.pdf" --delay 0.1
echo.

echo 测试完成！
pause






