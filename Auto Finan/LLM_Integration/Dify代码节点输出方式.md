# Dify 代码节点输出方式

## ⚠️ 常见错误

### 错误 1：'return' outside function

**错误代码**：
```python
# ❌ 错误：在模块级别使用 return
return {
    "success": True,
    "data": "result"
}
```

**错误原因**：
- Dify 代码节点的代码在模块级别执行
- 不能使用 `return` 语句（只能在函数内使用）

---

## ✅ 正确方式

### 方式 1：使用 `output` 变量（推荐）

```python
# ✅ 正确：使用 output 变量
output = {
    "success": True,
    "mcp_prompt": mcp_prompt
}
```

**说明**：
- Dify 会自动将 `output` 变量作为节点输出
- 这是最推荐的方式

---

### 方式 2：最后一行作为输出

```python
# ✅ 正确：最后一行作为输出
{
    "success": True,
    "mcp_prompt": mcp_prompt
}
```

**说明**：
- 代码的最后一行会自动作为输出
- 但这种方式不够清晰，不推荐

---

## 📝 完整示例

### 示例 1：Excel → 提示词节点

```python
import sys
import os

# 添加路径
sys.path.insert(0, r'C:\Users\FH\source\repos\Auto Finan\Auto Finan\LLM_Integration')

from workflow_core import process_excel_to_mcp_direct

# 获取输入
excel_path = inputs.get('excel_path', '')
sheet_name = inputs.get('sheet_name', '')
serial = inputs.get('serial', '')

# 验证参数
if not excel_path or not sheet_name or not serial:
    output = {
        "success": False,
        "error": "缺少必要参数"
    }
elif not os.path.exists(excel_path):
    output = {
        "success": False,
        "error": f"Excel 文件不存在: {excel_path}"
    }
else:
    try:
        # 生成提示词
        mcp_prompt = process_excel_to_mcp_direct(excel_path, sheet_name, serial)
        
        if not mcp_prompt:
            output = {
                "success": False,
                "error": "未能生成提示词"
            }
        else:
            output = {
                "success": True,
                "mcp_prompt": mcp_prompt
            }
    except Exception as e:
        output = {
            "success": False,
            "error": str(e)
        }
```

---

### 示例 2：处理 HTTP 响应节点

```python
import json

# 获取 HTTP 响应
http_response = inputs.get('http_response', {})

# 解析 JSON（如果是字符串）
if isinstance(http_response, str):
    try:
        http_response = json.loads(http_response)
    except:
        output = {
            "success": False,
            "error": "响应格式错误"
        }
    else:
        # 处理响应
        status = http_response.get("status", "unknown")
        output = {
            "success": status == "success",
            "status": status,
            "message": http_response.get("message", ""),
            "logs": http_response.get("logs", [])
        }
else:
    # 如果已经是字典
    status = http_response.get("status", "unknown")
    output = {
        "success": status == "success",
        "status": status,
        "message": http_response.get("message", ""),
        "logs": http_response.get("logs", [])
    }
```

---

## 🔍 语法对比

| 方式 | 语法 | 说明 |
|------|------|------|
| ❌ **错误** | `return {...}` | 不能在模块级别使用 |
| ✅ **正确** | `output = {...}` | 推荐方式 |
| ✅ **正确** | `{...}` (最后一行) | 可用但不推荐 |

---

## 💡 最佳实践

1. **始终使用 `output` 变量**：
   - 代码更清晰
   - 易于调试
   - 符合 Dify 规范

2. **使用 if-else 结构**：
   - 避免多个 return
   - 确保只有一个 output

3. **添加错误处理**：
   - 捕获异常
   - 返回错误信息

---

## 🐛 常见问题

### 问题 1：多个 output 赋值

```python
# ⚠️ 可能的问题：多个 output
if condition:
    output = {"success": True}
else:
    output = {"success": False}
# 这样是可以的，因为只有一个 output 会被执行
```

### 问题 2：忘记设置 output

```python
# ❌ 错误：没有设置 output
if condition:
    result = "success"
# 节点没有输出！
```

**解决**：
```python
# ✅ 正确：确保设置 output
if condition:
    output = {"result": "success"}
else:
    output = {"result": "failure"}
```

---

## 📚 参考

- Dify 官方文档：代码节点输出
- 完整配置指南：`Dify工作流配置指南.md`

