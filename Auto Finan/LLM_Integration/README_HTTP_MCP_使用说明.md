# Playwright MCP HTTP 调用使用说明

## 📋 问题说明

当你运行 `http_mcp_example.py` 时，如果遇到以下错误：

```
❌ 无法连接到 MCP 端点: http://localhost:3030/mcp/execute
```

这说明 **Playwright MCP HTTP 服务没有运行**。

---

## 🔧 解决方案

### 方案 1：使用 HTTP 网关（推荐）

我创建了一个 HTTP 网关服务，可以接收你的 HTTP 请求并解析提示词。

#### 步骤 1：启动 HTTP 网关

```bash
cd "Auto Finan\LLM_Integration"
start_mcp_gateway.bat
```

或者直接运行：

```bash
python playwright_mcp_http_gateway.py
```

网关将在 `http://localhost:3030` 启动。

#### 步骤 2：测试连接

```bash
# 在新的命令行窗口运行
python http_mcp_example.py
```

选择选项 1-4 进行测试。

#### 步骤 3：在 Dify 中使用

在 Dify 工作流中配置 HTTP 请求节点：
- **URL**: `http://localhost:3030/mcp/execute`
- **方法**: `POST`
- **请求体**:
```json
{
  "prompt": "{{#workflow.prompt#}}"
}
```

---

### 方案 2：通过 Cursor MCP 客户端执行（实际执行）

**重要**：Playwright MCP 主要通过 **Cursor 的 MCP 客户端** 来执行，而不是独立的 HTTP 服务。

#### 方式 A：在 Cursor 中直接使用

1. 确保你的 `mcp.json` 配置了 Playwright MCP：
```json
{
  "mcpServers": {
    "playwright": {
      "command": "npx",
      "args": ["@playwright/mcp@0.0.46"]
    }
  }
}
```

2. 在 Cursor 中直接发送提示词给 AI，AI 会自动调用 Playwright MCP 执行。

#### 方式 B：使用 HTTP 网关 + Cursor MCP

HTTP 网关可以：
- ✅ 接收 HTTP 请求
- ✅ 解析和验证提示词
- ✅ 返回格式化的响应

但**实际执行**需要通过 Cursor 的 MCP 客户端。

---

## 🚀 快速开始

### 1. 启动 HTTP 网关

```bash
cd "Auto Finan\LLM_Integration"
start_mcp_gateway.bat
```

### 2. 运行示例脚本

```bash
# 在新的命令行窗口
python http_mcp_example.py
```

### 3. 测试 HTTP 接口

使用 cURL 或 PowerShell：

```powershell
$body = @{
    prompt = "请你调用Playwright MCP，执行以下命令，一次性执行完`n打开https://example.com"
} | ConvertTo-Json

Invoke-RestMethod -Uri "http://localhost:3030/mcp/execute" -Method Post -Body $body -ContentType "application/json"
```

---

## 📝 提示词格式

HTTP 网关接收的提示词格式：

```json
{
  "prompt": "1. 请你调用Playwright MCP，执行以下命令，一次性执行完\n2. 打开https://example.com\n3. 在用户名输入框中输入test"
}
```

---

## ⚠️ 注意事项

1. **HTTP 网关的作用**：
   - ✅ 接收和解析提示词
   - ✅ 验证格式
   - ✅ 返回结构化响应
   - ❌ **不直接执行浏览器操作**

2. **实际执行方式**：
   - 通过 Cursor 的 MCP 客户端（推荐）
   - 或通过 Playwright MCP 的 SSE 接口（需要额外实现）

3. **端口冲突**：
   - 默认端口：`3030`
   - 如果端口被占用，修改 `playwright_mcp_http_gateway.py` 中的 `GATEWAY_PORT`

---

## 🔍 故障排查

### 问题 1：无法连接到端点

**原因**：HTTP 网关服务未启动

**解决**：
```bash
# 检查服务是否运行
curl http://localhost:3030/health

# 如果失败，启动服务
start_mcp_gateway.bat
```

### 问题 2：端口被占用

**解决**：
```bash
# 修改端口（在 playwright_mcp_http_gateway.py 中）
GATEWAY_PORT = 3031  # 改为其他端口

# 或设置环境变量
set GATEWAY_PORT=3031
```

### 问题 3：依赖缺失

**解决**：
```bash
pip install fastapi uvicorn
```

---

## 📚 相关文件

- `playwright_mcp_http_gateway.py` - HTTP 网关服务
- `start_mcp_gateway.bat` - 启动脚本
- `http_mcp_example.py` - 示例脚本
- `HTTP_MCP_调用示例.md` - 详细文档

---

## 💡 下一步

1. ✅ 启动 HTTP 网关服务
2. ✅ 运行示例脚本测试
3. ✅ 在 Dify 工作流中配置 HTTP 请求节点
4. 🔄 实际执行需要通过 Cursor MCP 客户端或实现 SSE 客户端

