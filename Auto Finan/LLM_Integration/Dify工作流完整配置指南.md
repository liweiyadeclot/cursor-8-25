# Dify 工作流完整配置指南

## ✅ 前提条件

- ✅ HTTP 请求已成功（本地服务运行正常）
- ✅ 本地服务地址：`http://192.168.137.133:8001`（或你的实际 IP）

---

## 🏗️ 工作流架构

```
┌─────────────┐
│  开始节点   │ 输入：excel_path, sheet_name, serial
└──────┬──────┘
       │
       ▼
┌─────────────────────────┐
│  HTTP请求节点（本地服务）│ 调用 /api/excel-to-prompt
│  生成 MCP 提示词        │
└──────┬──────────────────┘
       │
       ▼
┌─────────────┐
│  代码节点   │ 处理响应，提取 mcp_prompt
└──────┬──────┘
       │
       ▼
┌─────────────┐
│  条件判断   │ 检查是否成功
└──────┬──────┘
       │
       ├─ 成功 ──> HTTP请求（Playwright MCP）
       │
       └─ 失败 ──> 错误处理
```

---

## 📝 详细配置步骤

### 步骤 1：创建开始节点

**节点类型**：`开始`

**输入变量**（添加 3 个变量）：

| 变量名 | 类型 | 说明 | 示例值 |
|--------|------|------|--------|
| `excel_path` | 字符串 | Excel 文件路径 | `C:\Users\FH\source\repos\Auto Finan\Auto Finan\420财务050823.xlsx` |
| `sheet_name` | 字符串 | 工作表名称 | `3-报销` |
| `serial` | 字符串 | 序号 | `1` |

---

### 步骤 2：添加 HTTP 请求节点（调用本地服务）

**节点类型**：`HTTP 请求`

**配置项**：

| 配置项 | 值 |
|--------|-----|
| **方法** | `POST` |
| **URL** | `http://192.168.137.133:8001/api/excel-to-prompt`<br>（替换为你的实际 IP） |
| **请求头** | 见下方 |
| **请求体** | 见下方 |
| **超时** | `60` 秒 |

#### 请求头配置

**格式**：JSON

```json
{
  "Content-Type": "application/json"
}
```

#### 请求体配置

**格式**：JSON

```json
{
  "excel_path": "{{#workflow.excel_path#}}",
  "sheet_name": "{{#workflow.sheet_name#}}",
  "serial": "{{#workflow.serial#}}"
}
```

**重要提示**：
- 使用 `{{#workflow.变量名#}}` 引用工作流变量
- 确保 JSON 格式正确（双引号，无尾逗号）

#### 输出变量

HTTP 请求节点会自动创建输出变量：
- `http_response` - HTTP 响应内容（通常是 JSON 字符串）

---

### 步骤 3：添加代码节点（处理响应）

**节点类型**：`代码`

**输入变量配置**：
- `http_response` (string) - 来自 HTTP 请求节点

**代码**（官方标准方式）：

```python
def main(http_response: str) -> dict:
    """
    处理 HTTP 响应
    
    参数：
        http_response: HTTP 响应内容（JSON 字符串）
    """
    import json
    
    # 解析 JSON
    try:
        response = json.loads(http_response)
    except json.JSONDecodeError as e:
        return {
            "success": False,
            "error": f"JSON 解析失败: {str(e)}",
            "raw_response": http_response[:200] if len(http_response) > 200 else http_response
        }
    
    # 检查响应
    if not response:
        return {
            "success": False,
            "error": "响应为空"
        }
    
    # 处理响应
    if response.get("success"):
        # 成功：提取 MCP 提示词
        return {
            "success": True,
            "mcp_prompt": response.get("mcp_prompt", ""),
            "prompt_length": response.get("prompt_length", 0),
            "message": response.get("message", "提示词生成成功")
        }
    else:
        # 失败：提取错误信息
        return {
            "success": False,
            "error": response.get("error", "未知错误"),
            "debug": response.get("debug", {})
        }
```

**或者使用兼容版本**（如果官方方式不工作）：

```python
import json

# 获取 HTTP 响应
http_response = inputs.get('http_response', '')

# 解析 JSON
if isinstance(http_response, str):
    response = json.loads(http_response)
else:
    response = http_response

# 处理响应
if response.get("success"):
    output = {
        "success": True,
        "mcp_prompt": response.get("mcp_prompt", "")
    }
else:
    output = {
        "success": False,
        "error": response.get("error", "未知错误")
    }
```

**输出变量配置**（在 Dify 代码节点中配置）：

| 变量名 | 类型 | 说明 |
|--------|------|------|
| `success` | boolean | 是否成功 |
| `mcp_prompt` | string | MCP 提示词 |
| `prompt_length` | number | 提示词长度 |
| `error` | string | 错误信息（可选） |

**重要**：确保代码返回的字典包含所有这些变量，即使值为空！

---

### 步骤 4：添加条件判断节点

**节点类型**：`条件判断`

**条件表达式**：

```
{{#workflow.success#}} == true
```

**分支**：
- **True 分支**：继续执行（调用 Playwright MCP）
- **False 分支**：错误处理（可选）

---

### 步骤 5：添加 HTTP 请求节点（调用 Playwright MCP）

**节点类型**：`HTTP 请求`

**配置项**：

| 配置项 | 值 |
|--------|-----|
| **方法** | `POST` |
| **URL** | `http://localhost:3030/mcp/execute`<br>（或你的 MCP 服务地址） |
| **请求头** | 见下方 |
| **请求体** | 见下方 |
| **超时** | `300` 秒（5分钟，因为浏览器操作可能较慢） |

#### 请求头配置

```json
{
  "Content-Type": "application/json"
}
```

#### 请求体配置

```json
{
  "prompt": "{{#workflow.mcp_prompt#}}"
}
```

**重要**：
- 确保 Playwright MCP HTTP 网关服务正在运行
- 如果 Dify 在远程，需要将 `localhost` 替换为实际 IP

---

### 步骤 6：添加结果处理节点（可选）

**节点类型**：`代码`

**代码**：

```python
import json

# 获取 Playwright MCP 响应
mcp_response = inputs.get('http_response', '')

# 解析 JSON
if isinstance(mcp_response, str):
    try:
        response = json.loads(mcp_response)
    except:
        response = {"status": "unknown", "message": mcp_response}
else:
    response = mcp_response

# 处理结果
output = {
    "success": response.get("status") == "success",
    "status": response.get("status", "unknown"),
    "message": response.get("message", ""),
    "execution_id": response.get("execution_id", ""),
    "logs": response.get("logs", []),
    "full_response": response
}
```

---

### 步骤 7：添加结束节点

**节点类型**：`结束`

**输出变量**（可选）：
- 可以输出最终结果给调用方

---

## 🔧 关键配置要点

### 1. 变量引用格式

**在 HTTP 请求节点中**：
```json
{
  "excel_path": "{{#workflow.excel_path#}}"
}
```

**在代码节点中**：
```python
excel_path = inputs.get('excel_path', '')
```

**在条件判断中**：
```
{{#workflow.success#}} == true
```

---

### 2. IP 地址配置

**本地服务 IP**：
- 如果 Dify 在本地：`http://localhost:8001`
- 如果 Dify 在远程：`http://192.168.137.133:8001`（你的实际 IP）

**Playwright MCP IP**：
- 如果 Dify 在本地：`http://localhost:3030`
- 如果 Dify 在远程：`http://192.168.137.133:3030`（你的实际 IP）

---

### 3. 路径处理（推荐）

在代码节点中预处理路径，使用正斜杠：

```python
excel_path = inputs.get('excel_path', '').replace('\\', '/')
```

这样可以避免 JSON 转义问题。

---

## 📊 完整工作流示例

### 节点连接顺序

```
开始节点
  │
  ├─> HTTP请求（本地服务）
  │     │
  │     └─> 代码节点（处理响应）
  │           │
  │           └─> 条件判断
  │                 │
  │                 ├─ True ──> HTTP请求（Playwright MCP）
  │                 │            │
  │                 │            └─> 代码节点（结果处理）
  │                 │                  │
  │                 │                  └─> 结束节点
  │                 │
  │                 └─ False ──> 代码节点（错误处理）
  │                               │
  │                               └─> 结束节点
```

---

## 🧪 测试工作流

### 测试步骤

1. **设置输入变量**：
   - `excel_path`: `C:\Users\FH\source\repos\Auto Finan\Auto Finan\420财务050823.xlsx`
   - `sheet_name`: `3-报销`
   - `serial`: `1`

2. **运行工作流**

3. **检查每个节点的输出**：
   - HTTP 请求节点：查看响应状态码和内容
   - 代码节点：查看 `output` 变量
   - 条件判断：查看分支选择

---

## ⚠️ 常见问题

### 问题 1：HTTP 请求失败

**检查**：
- ✅ 本地服务是否运行
- ✅ IP 地址是否正确
- ✅ 防火墙是否允许连接

### 问题 2：JSON 解析失败

**检查**：
- ✅ 响应格式是否正确
- ✅ 响应是否为有效的 JSON

### 问题 3：变量未定义

**检查**：
- ✅ 变量名是否正确
- ✅ 变量是否在上游节点中定义

---

## 📚 相关文档

- `Dify_HTTP请求体配置.md` - HTTP 请求详细配置
- `Dify使用本地服务配置.md` - 本地服务配置
- `Dify工作流快速配置.md` - 快速配置指南

---

## ✅ 配置检查清单

- [ ] 开始节点：已添加 3 个输入变量
- [ ] HTTP 请求节点（本地服务）：URL、请求头、请求体已配置
- [ ] 代码节点（处理响应）：代码已添加
- [ ] 条件判断节点：条件表达式已设置
- [ ] HTTP 请求节点（Playwright MCP）：URL、请求体已配置
- [ ] 本地服务正在运行
- [ ] Playwright MCP 服务正在运行（如果需要）
- [ ] 网络连接正常

---

## 🎉 完成！

配置完成后，你就可以在 Dify 中运行工作流，自动将 Excel 数据转换为 MCP 提示词并执行了！

