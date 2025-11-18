@echo off
setlocal

REM 启动 Playwright MCP HTTP 服务器
REM 注意：由于 0.0.47 版本存在 utilsBundleImpl 错误，使用 0.0.46 版本

echo ================================================================================
echo 启动 Playwright MCP HTTP 服务器
echo ================================================================================
echo.

REM 检查 Node.js
where node >nul 2>nul
if %errorlevel% neq 0 (
  echo ❌ 未找到 Node.js，请先安装 Node.js
  echo 💡 下载地址: https://nodejs.org/
  pause
  exit /b 1
)

echo ✅ Node.js 已安装
node -v
echo.

REM 设置服务器参数
set MCP_PORT=3030
set MCP_HOST=0.0.0.0

echo 📡 启动参数:
echo    端口: %MCP_PORT%
echo    主机: %MCP_HOST%
echo    版本: @playwright/mcp@0.0.46 (使用稳定版本)
echo.

echo 🚀 正在启动 Playwright MCP HTTP 服务器...
echo 💡 提示：服务器启动后，请保持此窗口打开
echo 💡 访问地址: http://localhost:%MCP_PORT%
echo.

REM 启动 Playwright MCP HTTP 服务器
REM 使用 0.0.46 版本避免 utilsBundleImpl 错误
npx @playwright/mcp@0.0.46 --port %MCP_PORT% --host %MCP_HOST% --headless

if %errorlevel% neq 0 (
  echo.
  echo ❌ 启动失败
  echo 💡 可能的原因:
  echo    1. 端口 %MCP_PORT% 已被占用
  echo    2. 网络连接问题
  echo    3. npm 缓存问题（尝试运行: npm cache clean --force）
  pause
  exit /b 1
)

pause

