@echo off
echo 启动Qwen 2.5 7B报销信息提取测试...
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: Python未安装或未添加到PATH
    pause
    exit /b 1
)

REM 检查Ollama服务是否运行
curl -s http://localhost:11434/api/tags >nul 2>&1
if errorlevel 1 (
    echo 警告: Ollama服务可能未运行
    echo 请确保已运行: ollama serve
    echo.
)

REM 安装依赖
echo 检查Python依赖...
pip install -r requirements.txt

REM 运行测试
echo 开始测试...
echo.
python test_qwen_extraction.py

pause
