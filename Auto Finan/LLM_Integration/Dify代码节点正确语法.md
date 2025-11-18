# Dify 代码节点变量访问语法

## ⚠️ 常见错误

### 错误示例（会导致语法错误）

```python
# ❌ 错误：在代码节点中不能使用 {{#workflow.variable_name#}}
excel_path = {{#workflow.excel_path#}}  # SyntaxError!
```

**错误原因**：
- `{{#workflow.variable_name#}}` 是模板语法，用于 HTTP 请求节点等
- 在代码节点中会被当作 Python 代码，导致语法错误

---

## ✅ 正确语法

### 在代码节点中访问变量

**使用 `inputs` 字典**：

```python
# ✅ 正确：使用 inputs 字典
excel_path = inputs.get('excel_path', '')
sheet_name = inputs.get('sheet_name', '')
serial = inputs.get('serial', '')
```

**或者直接访问**：

```python
# ✅ 也可以直接访问（如果确定变量存在）
excel_path = inputs['excel_path']
sheet_name = inputs['sheet_name']
serial = inputs['serial']
```

---

## 📝 完整示例

### 节点 1：Excel → 提示词（代码节点）

```python
import sys
import os

# 添加路径
sys.path.insert(0, r'C:\Users\FH\source\repos\Auto Finan\Auto Finan\LLM_Integration')

from workflow_core import process_excel_to_mcp_direct

# ✅ 正确：使用 inputs 字典
excel_path = inputs.get('excel_path', '')
sheet_name = inputs.get('sheet_name', '')
serial = inputs.get('serial', '')

# 验证参数
if not excel_path or not sheet_name or not serial:
    return {
        "success": False,
        "error": "缺少必要参数"
    }

# 生成提示词
try:
    mcp_prompt = process_excel_to_mcp_direct(excel_path, sheet_name, serial)
    
    return {
        "success": True,
        "mcp_prompt": mcp_prompt
    }
except Exception as e:
    return {
        "success": False,
        "error": str(e)
    }
```

### 节点 2：HTTP 请求节点

**在 HTTP 请求节点中，可以使用模板语法**：

- **URL**：`http://localhost:3030/mcp/execute`
- **请求体**：
```json
{
  "prompt": "{{#workflow.mcp_prompt#}}"
}
```

**注意**：HTTP 请求节点可以使用 `{{#workflow.variable_name#}}` 语法。

### 节点 3：处理 HTTP 响应（代码节点）

```python
import json

# ✅ 正确：使用 inputs 字典
http_response = inputs.get('http_response', {})

# 如果是字符串，解析为 JSON
if isinstance(http_response, str):
    try:
        http_response = json.loads(http_response)
    except:
        return {
            "success": False,
            "error": "响应格式错误"
        }

# 处理响应
return {
    "status": http_response.get("status", "unknown"),
    "message": http_response.get("message", ""),
    "logs": http_response.get("logs", [])
}
```

---

## 🔍 语法对比表

| 节点类型 | 变量访问方式 | 示例 |
|---------|------------|------|
| **代码节点** | `inputs.get('variable_name')` | `excel_path = inputs.get('excel_path', '')` |
| **HTTP 请求节点** | `{{#workflow.variable_name#}}` | `"prompt": "{{#workflow.mcp_prompt#}}"` |
| **条件判断节点** | `{{#workflow.variable_name#}}` | `{{#workflow.success#}} == true` |

---

## 💡 重要提示

1. **代码节点**：必须使用 `inputs` 字典
2. **HTTP 请求节点**：使用 `{{#workflow.variable_name#}}` 模板语法
3. **条件判断节点**：使用 `{{#workflow.variable_name#}}` 模板语法

---

## 🐛 常见错误和解决方案

### 错误 1：SyntaxError: '{' was never closed

**原因**：在代码节点中使用了 `{{#workflow.variable_name#}}`

**解决**：改为 `inputs.get('variable_name', '')`

### 错误 2：KeyError: 'variable_name'

**原因**：变量不存在

**解决**：使用 `inputs.get('variable_name', '')` 并提供默认值

### 错误 3：变量值为空

**原因**：变量未正确传递

**解决**：
1. 检查上游节点是否正确输出变量
2. 检查变量名是否匹配
3. 添加调试输出：`return {"debug": inputs}`

---

## 📚 参考

- Dify 官方文档：代码节点变量访问
- 完整配置指南：`Dify工作流配置指南.md`

