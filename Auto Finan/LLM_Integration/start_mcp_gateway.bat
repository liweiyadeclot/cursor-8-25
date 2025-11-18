@echo off
setlocal

REM 启动 Playwright MCP HTTP 网关服务

echo ================================================================================
echo 启动 Playwright MCP HTTP 网关服务
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

REM 设置环境变量
set GATEWAY_PORT=3030
set GATEWAY_HOST=0.0.0.0

echo 📡 启动参数:
echo    端口: %GATEWAY_PORT%
echo    主机: %GATEWAY_HOST%
echo.

echo 🚀 正在启动 HTTP 网关服务...
echo 💡 提示：服务器启动后，请保持此窗口打开
echo 💡 访问地址: http://localhost:%GATEWAY_PORT%
echo 💡 执行端点: http://localhost:%GATEWAY_PORT%/mcp/execute
echo 💡 健康检查: http://localhost:%GATEWAY_PORT%/health
echo.

REM 切换到脚本所在目录
cd /d "%~dp0"

REM 启动服务（使用真正执行版本）
python playwright_mcp_http_gateway_executor.py

if %errorlevel% neq 0 (
  echo.
  echo ❌ 启动失败
  pause
  exit /b 1
)

pause

