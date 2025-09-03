@echo off
chcp 65001 >nul
echo 综合功能测试
echo =============

echo.
echo 1. 测试基础功能...
python simple_python_test.py

echo.
echo 2. 测试文件夹创建功能...
python test_folder_creation.py

echo.
echo 3. 测试Python脚本帮助信息...
python test_mouse_keyboard.py --help

echo.
echo 4. 测试配置文件读取...
python -c "import json; config=json.load(open('config.json', 'r', encoding='utf-8')); print('配置文件读取成功'); print('坐标配置:', list(config['ScreenPositions'].keys()))"

echo.
echo 5. 测试Enter键功能（可选）...
echo 注意：这个测试需要手动操作，请确保焦点在正确的输入框中
echo 如果要测试Enter键功能，请运行: test_enter_key.bat

echo.
echo 所有测试完成！
echo.
echo 如果要进行实际的功能测试，请运行：
echo test_python_script_only.bat
echo.
pause
