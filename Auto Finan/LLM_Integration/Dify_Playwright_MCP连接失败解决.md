# Dify Playwright MCP 连接失败解决

## ❌ 错误信息

```
Reached maximum retries (0) for URL http://localhost:3030/mcp/execute
```

---

## 🔍 问题原因

**Playwright MCP HTTP 网关服务没有运行**，或者无法连接。

---

## ✅ 解决方案

### 步骤 1：启动 Playwright MCP HTTP 网关服务

在本地运行以下命令：

```bash
cd "Auto Finan\LLM_Integration"
start_mcp_gateway.bat
```

或者直接运行：

```bash
python playwright_mcp_http_gateway.py
```

服务将在 `http://localhost:3030` 启动。

---

### 步骤 2：验证服务运行

在浏览器或命令行中测试：

```bash
# 健康检查
curl http://localhost:3030/health

# 或使用 PowerShell
Invoke-RestMethod -Uri "http://localhost:3030/health" -Method Get
```

应该返回：
```json
{
  "status": "ok",
  "service": "Playwright MCP HTTP Gateway"
}
```

---

### 步骤 3：检查 Dify 配置

在 Dify HTTP 请求节点中：

**如果 Dify 在本地**：
- URL: `http://localhost:3030/mcp/execute`

**如果 Dify 在远程服务器**：
- URL: `http://你的本地IP:3030/mcp/execute`
- 例如：`http://192.168.137.133:3030/mcp/execute`

**重要**：
- 确保防火墙允许端口 3030
- 确保服务绑定到 `0.0.0.0` 而不是 `127.0.0.1`

---

## 🔧 详细步骤

### 1. 启动服务

**方法 A：使用批处理文件（推荐）**

```bash
cd "C:\Users\FH\source\repos\Auto Finan\Auto Finan\LLM_Integration"
start_mcp_gateway.bat
```

**方法 B：直接运行 Python**

```bash
cd "C:\Users\FH\source\repos\Auto Finan\Auto Finan\LLM_Integration"
python playwright_mcp_http_gateway.py
```

**方法 C：使用 uvicorn**

```bash
cd "C:\Users\FH\source\repos\Auto Finan\Auto Finan\LLM_Integration"
uvicorn playwright_mcp_http_gateway:app --host 0.0.0.0 --port 3030
```

---

### 2. 验证服务运行

服务启动后，应该看到类似输出：

```
INFO:     Started server process [...]
INFO:     Waiting for application startup.
INFO:     Application startup complete.
INFO:     Uvicorn running on http://0.0.0.0:3030
```

---

### 3. 测试连接

**在 PowerShell 中测试**：

```powershell
$body = @{
    prompt = "测试提示词"
} | ConvertTo-Json

Invoke-RestMethod -Uri "http://localhost:3030/mcp/execute" -Method Post -Body $body -ContentType "application/json"
```

---

## 🌐 网络配置

### 情况 1：Dify 在本地

**配置**：
- URL: `http://localhost:3030/mcp/execute`

**优点**：
- 无需网络配置
- 速度最快

---

### 情况 2：Dify 在远程服务器

**配置步骤**：

1. **获取本地 IP 地址**：
   ```bash
   # Windows
   ipconfig
   # 查找 IPv4 地址，例如：192.168.137.133
   ```

2. **配置防火墙**：
   - 允许端口 3030 入站
   - Windows 防火墙：添加端口 3030 的入站规则

3. **确保服务绑定到 0.0.0.0**：
   - 检查 `playwright_mcp_http_gateway.py` 中的 `GATEWAY_HOST`
   - 应该是 `0.0.0.0` 而不是 `127.0.0.1`

4. **在 Dify 中使用**：
   - URL: `http://192.168.137.133:3030/mcp/execute`
   - 替换为你的实际 IP 地址

---

## ⚠️ 常见问题

### 问题 1：端口被占用

**错误**：
```
Error: [Errno 48] Address already in use
```

**解决**：
1. 查找占用端口的进程：
   ```bash
   netstat -ano | findstr :3030
   ```
2. 结束进程或修改端口

---

### 问题 2：服务启动失败

**检查**：
1. Python 是否安装
2. 依赖是否安装：`pip install fastapi uvicorn`
3. 查看错误信息

---

### 问题 3：Dify 无法连接

**检查**：
1. 服务是否运行
2. IP 地址是否正确
3. 防火墙是否允许
4. 服务是否绑定到 `0.0.0.0`

---

## 📝 完整工作流检查清单

- [ ] Playwright MCP HTTP 网关服务正在运行
- [ ] 服务监听在 `0.0.0.0:3030`（不是 `127.0.0.1`）
- [ ] 健康检查通过：`http://localhost:3030/health`
- [ ] Dify HTTP 请求节点 URL 正确
- [ ] 如果 Dify 在远程，IP 地址和防火墙配置正确
- [ ] 请求体格式正确：`{"prompt": "{{#workflow.mcp_prompt#}}"}`

---

## 🚀 快速修复

### 一键启动服务

```bash
# 在 PowerShell 中
cd "C:\Users\FH\source\repos\Auto Finan\Auto Finan\LLM_Integration"
.\start_mcp_gateway.bat
```

### 验证服务

```bash
# 在新的 PowerShell 窗口
Invoke-RestMethod -Uri "http://localhost:3030/health" -Method Get
```

---

## 📚 相关文件

- `playwright_mcp_http_gateway.py` - HTTP 网关服务
- `start_mcp_gateway.bat` - 启动脚本
- `README_HTTP_MCP_使用说明.md` - 详细文档

---

## ✅ 总结

**问题**：无法连接到 `http://localhost:3030/mcp/execute`

**解决**：
1. ✅ 启动 Playwright MCP HTTP 网关服务
2. ✅ 验证服务运行（健康检查）
3. ✅ 检查 Dify 配置（URL、IP、防火墙）

**关键**：确保服务在运行，并且可以从 Dify 访问！

