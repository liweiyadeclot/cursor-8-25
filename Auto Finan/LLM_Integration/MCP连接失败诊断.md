# MCP 连接失败诊断指南

## 错误信息
```
Reached maximum retries (0) for URL http://192.168.137.133:3030/mcp/execute
```

## 可能的原因

### 1. 服务未启动
- **检查方法**：查看是否有 `start_mcp_gateway.bat` 的窗口正在运行
- **解决方案**：运行 `start_mcp_gateway.bat` 启动服务

### 2. 服务监听地址不正确
- **问题**：服务可能监听在 `localhost:3030` 而不是 `0.0.0.0:3030`
- **检查方法**：查看服务启动日志，确认显示 `http://0.0.0.0:3030`
- **解决方案**：确保服务监听在 `0.0.0.0`（允许远程连接）

### 3. Windows 防火墙阻止连接
- **检查方法**：
  1. 打开 Windows 防火墙设置
  2. 检查入站规则是否允许端口 3030
- **解决方案**：
  ```powershell
  # 以管理员身份运行 PowerShell，添加防火墙规则
  New-NetFirewallRule -DisplayName "MCP Gateway" -Direction Inbound -LocalPort 3030 -Protocol TCP -Action Allow
  ```

### 4. 端口被占用
- **检查方法**：
  ```powershell
  netstat -ano | findstr :3030
  ```
- **解决方案**：
  - 如果端口被占用，关闭占用端口的进程
  - 或者修改服务端口（通过环境变量 `GATEWAY_PORT`）

### 5. 网络连接问题
- **检查方法**：
  ```powershell
  # 从客户端测试连接
  Test-NetConnection -ComputerName 192.168.137.133 -Port 3030
  ```
- **解决方案**：确保客户端和服务器在同一网络，或网络路由正确

## 快速诊断步骤

### 步骤 1：检查服务是否运行
在服务端（192.168.137.133）运行：
```powershell
cd "C:\Users\FH\source\repos\Auto Finan\Auto Finan\LLM_Integration"
python "检查MCP服务状态.py" --host localhost
```

### 步骤 2：检查远程连接
在客户端运行：
```powershell
python "检查MCP服务状态.py" --url "http://192.168.137.133:3030"
```

或者使用 PowerShell：
```powershell
Test-NetConnection -ComputerName 192.168.137.133 -Port 3030
```

### 步骤 3：检查服务监听地址
查看服务启动日志，应该显示：
```
🌐 服务地址: http://0.0.0.0:3030
```

如果显示 `http://127.0.0.1:3030` 或 `http://localhost:3030`，则需要修改服务配置。

## 解决方案

### 方案 1：确保服务监听在 0.0.0.0
服务默认应该监听在 `0.0.0.0:3030`，这允许来自任何 IP 地址的连接。

检查 `playwright_mcp_http_gateway_executor.py` 中的配置：
```python
GATEWAY_HOST = os.environ.get("GATEWAY_HOST", "0.0.0.0")
GATEWAY_PORT = int(os.environ.get("GATEWAY_PORT", "3030"))
```

### 方案 2：添加防火墙规则
以管理员身份运行 PowerShell：
```powershell
New-NetFirewallRule -DisplayName "MCP Gateway Port 3030" -Direction Inbound -LocalPort 3030 -Protocol TCP -Action Allow
```

### 方案 3：使用 localhost（如果客户端和服务器在同一台机器）
如果 Dify 和 MCP 服务在同一台机器上，可以修改 Dify 工作流中的 URL 为：
```
http://localhost:3030/mcp/execute
```

### 方案 4：检查服务进程
```powershell
# 查看是否有 Python 进程在运行服务
Get-Process python | Where-Object {$_.Path -like "*python*"}
```

## 验证服务正常运行

服务正常运行时，应该能够访问：
- 健康检查：`http://192.168.137.133:3030/health`
- 执行端点：`http://192.168.137.133:3030/mcp/execute`

使用浏览器或 curl 测试：
```powershell
# 测试健康检查
Invoke-WebRequest -Uri "http://192.168.137.133:3030/health" -Method GET

# 测试执行端点（需要 POST）
$body = @{prompt="测试"} | ConvertTo-Json
Invoke-WebRequest -Uri "http://192.168.137.133:3030/mcp/execute" -Method POST -Body $body -ContentType "application/json"
```

## 常见问题

### Q: 为什么显示 "Reached maximum retries (0)"？
A: 这表示连接在第一次尝试时就失败了，没有进行重试。通常是因为：
- 服务未启动
- 端口未开放
- 防火墙阻止

### Q: 如何确认服务正在监听 0.0.0.0？
A: 查看服务启动时的日志输出，应该显示：
```
🌐 服务地址: http://0.0.0.0:3030
```

### Q: 服务启动后立即关闭怎么办？
A: 检查是否有错误日志，可能是：
- 端口被占用
- 缺少依赖包
- Python 版本不兼容

## 联系支持

如果以上方法都无法解决问题，请提供：
1. 服务启动日志
2. 错误信息完整内容
3. 网络配置信息（IP 地址、子网掩码等）

