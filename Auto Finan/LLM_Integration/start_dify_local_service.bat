@echo off
setlocal

REM 启动 Dify 本地服务

echo ================================================================================
echo 启动 Dify 本地服务
echo ================================================================================
echo.

REM 检查 Python
where python >nul 2>nul
if %errorlevel% neq 0 (
  echo ❌ 未找到 Python，请先安装 Python
  pause
  exit /b 1
)

echo ✅ Python 已安装
python --version
echo.

REM 检查依赖
echo 🔍 检查依赖...
python -c "import fastapi, uvicorn" 2>nul
if %errorlevel% neq 0 (
  echo ⚠️  缺少依赖，正在安装...
  python -m pip install fastapi uvicorn
  if %errorlevel% neq 0 (
    echo ❌ 依赖安装失败
    pause
    exit /b 1
  )
  echo ✅ 依赖安装完成
)

echo.

REM 设置服务参数
set SERVICE_PORT=8001
set SERVICE_HOST=0.0.0.0

echo 📡 启动参数:
echo    端口: %SERVICE_PORT%
echo    主机: %SERVICE_HOST%
echo.

echo 🚀 正在启动本地服务...
echo 💡 提示：服务启动后，请保持此窗口打开
echo 💡 访问地址: http://localhost:%SERVICE_PORT%
echo 💡 如果 Dify 在远程服务器，使用本机 IP 地址
echo.

REM 切换到脚本所在目录
cd /d "%~dp0"

REM 启动服务（使用灵活版本，避免 422 错误）
python dify_local_service_flexible.py --host %SERVICE_HOST% --port %SERVICE_PORT%

if %errorlevel% neq 0 (
  echo.
  echo ❌ 启动失败
  pause
  exit /b 1
)

pause

