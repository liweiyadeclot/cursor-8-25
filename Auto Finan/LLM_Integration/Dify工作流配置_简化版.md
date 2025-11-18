# Dify 工作流配置（简化版）

## 🚀 3 步完成配置

### 步骤 1：开始节点

**输入变量**：
- `excel_path` (字符串)
- `sheet_name` (字符串)
- `serial` (字符串)

---

### 步骤 2：HTTP 请求节点（调用本地服务）

**配置**：

| 项目 | 值 |
|------|-----|
| 方法 | `POST` |
| URL | `http://192.168.137.133:8001/api/excel-to-prompt` |
| 请求头 | `{"Content-Type": "application/json"}` |
| 请求体 | `{"excel_path": "{{#workflow.excel_path#}}", "sheet_name": "{{#workflow.sheet_name#}}", "serial": "{{#workflow.serial#}}"}` |

**输出**：`http_response`（包含 `mcp_prompt`）

---

### 步骤 3：代码节点（提取提示词）

**代码**：

```python
import json

response = json.loads(inputs.get('http_response', '{}'))
output = {
    "success": response.get("success", False),
    "mcp_prompt": response.get("mcp_prompt", "")
}
```

---

## 📝 后续步骤（可选）

### 调用 Playwright MCP

添加 HTTP 请求节点：

| 项目 | 值 |
|------|-----|
| 方法 | `POST` |
| URL | `http://localhost:3030/mcp/execute` |
| 请求体 | `{"prompt": "{{#workflow.mcp_prompt#}}"}` |

---

## ✅ 完成！

现在工作流可以：
1. 接收 Excel 路径、工作表、序号
2. 调用本地服务生成 MCP 提示词
3. （可选）调用 Playwright MCP 执行

