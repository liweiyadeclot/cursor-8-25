# Dify 使用本地服务配置指南

## 🎯 方案概述

**问题**：Dify 运行在服务器上，无法直接访问本地的 `workflow_core.py` 文件。

**解决方案**：在本地运行 HTTP 服务，Dify 通过 HTTP 请求调用本地功能。

---

## 🚀 快速开始

### 步骤 1：启动本地服务

在本地运行：

```bash
cd "Auto Finan\LLM_Integration"
start_dify_local_service.bat
```

或者：

```bash
python dify_local_service.py
```

服务将在 `http://localhost:8001` 启动。

---

### 步骤 2：配置 Dify 工作流

#### 2.1 添加开始节点

**输入变量**：
- `excel_path` - Excel 文件路径
- `sheet_name` - 工作表名称
- `serial` - 序号

---

#### 2.2 添加 HTTP 请求节点（调用本地服务）

**节点类型**：`HTTP 请求`

**配置**：

| 项目 | 值 |
|------|-----|
| 方法 | POST |
| URL | `http://你的本地IP:8001/api/excel-to-prompt` |
| 请求头 | `{"Content-Type": "application/json"}` |
| 请求体 | `{"excel_path": "{{#workflow.excel_path#}}", "sheet_name": "{{#workflow.sheet_name#}}", "serial": "{{#workflow.serial#}}"}` |
| 超时 | 60 秒 |

**重要**：
- 如果 Dify 在本地：使用 `http://localhost:8001`
- 如果 Dify 在远程服务器：使用 `http://你的本地IP:8001`
  - 例如：`http://192.168.1.100:8001`

---

#### 2.3 添加代码节点（处理响应）

**节点类型**：`代码`

**代码**：

```python
import json

# 获取 HTTP 响应
response = inputs.get('http_response', {})

# 解析 JSON（如果是字符串）
if isinstance(response, str):
    try:
        response = json.loads(response)
    except:
        output = {
            "success": False,
            "error": "响应格式错误"
        }

# 检查响应
if response.get("success"):
    output = {
        "success": True,
        "mcp_prompt": response.get("mcp_prompt", ""),
        "prompt_length": response.get("prompt_length", 0)
    }
else:
    output = {
        "success": False,
        "error": response.get("error", "未知错误")
    }
```

---

#### 2.4 添加 HTTP 请求节点（调用 Playwright MCP）

**节点类型**：`HTTP 请求`

**配置**：

| 项目 | 值 |
|------|-----|
| 方法 | POST |
| URL | `http://localhost:3030/mcp/execute` |
| 请求头 | `{"Content-Type": "application/json"}` |
| 请求体 | `{"prompt": "{{#workflow.mcp_prompt#}}"}` |

---

## 📊 完整工作流

```
开始节点
  │
  ├─> [HTTP请求：调用本地服务 /api/excel-to-prompt]
  │     │
  │     └─> [代码节点：处理响应]
  │           │
  │           └─> [HTTP请求：调用 Playwright MCP]
  │                 │
  │                 └─> [代码节点：处理最终结果]
  │                       │
  │                       └─> [结束节点]
```

---

## 🌐 网络配置

### 情况 1：Dify 在本地

**配置**：
- URL: `http://localhost:8001/api/excel-to-prompt`

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
   
   # 查找 IPv4 地址，例如：192.168.1.100
   ```

2. **配置防火墙**：
   - 允许端口 8001 入站
   - Windows 防火墙：添加端口 8001 的入站规则

3. **在 Dify 中使用**：
   - URL: `http://192.168.1.100:8001/api/excel-to-prompt`
   - 替换为你的实际 IP 地址

4. **测试连接**：
   ```bash
   # 在 Dify 服务器上测试
   curl http://192.168.1.100:8001/health
   ```

---

### 情况 3：使用内网穿透（推荐用于生产环境）

如果需要从外网访问，可以使用：

1. **ngrok**：
   ```bash
   ngrok http 8001
   # 会生成一个公网 URL，例如：https://abc123.ngrok.io
   ```

2. **frp**：
   - 配置 frp 客户端和服务器
   - 映射本地 8001 端口

3. **在 Dify 中使用**：
   - URL: `https://abc123.ngrok.io/api/excel-to-prompt`

---

## 📝 API 接口说明

### 1. Excel 转提示词

**端点**：`POST /api/excel-to-prompt`

**请求体**：
```json
{
  "excel_path": "C:\\Users\\FH\\source\\repos\\Auto Finan\\Auto Finan\\420财务050823.xlsx",
  "sheet_name": "3-报销",
  "serial": "1"
}
```

**响应**：
```json
{
  "success": true,
  "mcp_prompt": "1. 请你调用Playwright MCP...",
  "prompt_length": 877,
  "excel_path": "...",
  "sheet_name": "3-报销",
  "serial": "1"
}
```

---

### 2. 批量处理

**端点**：`POST /api/batch-process`

**请求体**：
```json
{
  "excel_path": "C:\\Users\\FH\\source\\repos\\Auto Finan\\Auto Finan\\420财务050823.xlsx",
  "sheet_name": "3-报销"
}
```

**响应**：
```json
{
  "success": true,
  "results": [
    {
      "serial": "1",
      "json_data": {...},
      "mcp_prompt": "..."
    },
    ...
  ],
  "count": 3
}
```

---

### 3. 健康检查

**端点**：`GET /health`

**响应**：
```json
{
  "status": "healthy",
  "service": "Dify Local Service"
}
```

---

## 🔧 使用示例

### 在 Dify 中测试

**输入参数**：
```json
{
  "excel_path": "C:\\Users\\FH\\source\\repos\\Auto Finan\\Auto Finan\\420财务050823.xlsx",
  "sheet_name": "3-报销",
  "serial": "1"
}
```

**预期流程**：
1. Dify 发送 HTTP 请求到本地服务
2. 本地服务调用 `workflow_core.py` 处理
3. 返回 MCP 提示词
4. Dify 继续后续流程

---

## ⚠️ 注意事项

1. **文件路径**：
   - Excel 文件路径必须是本地服务可以访问的
   - 如果 Dify 在远程，需要确保文件在本地服务器上

2. **服务保持运行**：
   - 本地服务需要一直运行
   - 可以设置为 Windows 服务或使用进程管理器

3. **安全性**：
   - 生产环境建议添加认证
   - 使用 HTTPS（通过反向代理）

4. **性能**：
   - 本地服务处理速度很快
   - 网络延迟取决于 Dify 和本地服务的距离

---

## 🆚 方案对比

| 方案 | 优点 | 缺点 |
|------|------|------|
| **内联代码** | 无需额外服务 | 代码冗长，难以维护 |
| **本地 HTTP 服务** | 代码清晰，易于维护 | 需要运行额外服务 |

**推荐**：使用本地 HTTP 服务方案，代码更清晰，易于维护。

---

## 📚 相关文件

- `dify_local_service.py` - 本地服务代码
- `start_dify_local_service.bat` - 启动脚本
- `workflow_core.py` - 核心处理逻辑（本地）

---

## 💡 下一步

1. ✅ 启动本地服务：`start_dify_local_service.bat`
2. ✅ 配置 Dify HTTP 请求节点
3. ✅ 测试连接
4. ✅ 开始使用

