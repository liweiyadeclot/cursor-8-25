# Dify KeyError 修复说明

## ❌ 错误信息

```
KeyError: 'data'
```

**错误原因**：
- 代码试图访问 `data['data']['name']`
- 但实际响应格式中没有 `data` 键

---

## 🔍 问题分析

### 官方文档示例

官方文档示例：
```python
def main(http_response: str) -> dict:
    import json
    data = json.loads(http_response)
    return {
        'result': data['data']['name']  # 假设响应格式是 {"data": {"name": "..."}}
    }
```

### 我们的实际响应格式

我们的服务返回：
```json
{
  "success": true,
  "mcp_prompt": "...",
  "prompt_length": 123
}
```

**没有 `data` 键！**

---

## ✅ 修复方案

### 修复后的代码

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
            "error": f"JSON 解析失败: {str(e)}"
        }
    
    # 检查响应格式
    if not isinstance(response, dict):
        return {
            "success": False,
            "error": f"响应格式错误，期望字典，得到: {type(response).__name__}"
        }
    
    # 处理响应（根据实际响应格式）
    if response.get("success"):
        # 成功：提取 MCP 提示词
        return {
            "success": True,
            "mcp_prompt": response.get("mcp_prompt", ""),
            "prompt_length": response.get("prompt_length", 0)
        }
    else:
        # 失败：提取错误信息
        return {
            "success": False,
            "error": response.get("error", "未知错误")
        }
```

---

## 🔧 关键修改

### 1. 使用 `.get()` 方法

**错误**：
```python
# ❌ 直接访问，如果键不存在会报 KeyError
result = data['data']['name']
```

**正确**：
```python
# ✅ 使用 .get() 方法，提供默认值
mcp_prompt = response.get("mcp_prompt", "")
```

### 2. 根据实际响应格式访问

**我们的响应格式**：
```json
{
  "success": true,
  "mcp_prompt": "...",
  "prompt_length": 123
}
```

**访问方式**：
```python
if response.get("success"):
    mcp_prompt = response.get("mcp_prompt", "")
    prompt_length = response.get("prompt_length", 0)
```

---

## 📝 完整代码（直接使用）

```python
def main(http_response: str) -> dict:
    """
    处理 HTTP 响应
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
    
    # 检查响应格式
    if not isinstance(response, dict):
        return {
            "success": False,
            "error": f"响应格式错误，期望字典，得到: {type(response).__name__}"
        }
    
    # 处理响应
    if response.get("success"):
        return {
            "success": True,
            "mcp_prompt": response.get("mcp_prompt", ""),
            "prompt_length": response.get("prompt_length", 0),
            "message": response.get("message", "提示词生成成功")
        }
    else:
        return {
            "success": False,
            "error": response.get("error", "未知错误"),
            "debug": response.get("debug", {})
        }
```

---

## 🔍 调试：查看实际响应格式

如果不确定响应格式，可以使用以下调试代码：

```python
def main(http_response: str) -> dict:
    import json
    
    try:
        response = json.loads(http_response)
    except:
        return {
            "debug": {
                "raw_response": http_response[:500],
                "error": "JSON 解析失败"
            }
        }
    
    # 返回完整的响应结构用于调试
    return {
        "debug": {
            "response_type": type(response).__name__,
            "response_keys": list(response.keys()) if isinstance(response, dict) else [],
            "full_response": response
        }
    }
```

运行后查看输出，了解实际的响应格式。

---

## ⚠️ 常见错误

### 错误 1：KeyError: 'data'

**原因**：假设响应有 `data` 键，但实际没有

**解决**：使用 `.get()` 方法，或根据实际响应格式访问

---

### 错误 2：假设响应格式

**原因**：使用官方示例的格式，但实际服务返回不同格式

**解决**：
1. 先运行调试代码查看实际格式
2. 根据实际格式修改代码

---

### 错误 3：未检查响应类型

**原因**：直接访问字典键，但响应可能是字符串或其他类型

**解决**：先检查类型，再访问

```python
if isinstance(response, dict):
    value = response.get("key", "")
else:
    # 处理其他类型
    pass
```

---

## ✅ 最佳实践

1. **使用 `.get()` 方法**：避免 KeyError
2. **检查响应格式**：确保是期望的类型
3. **提供默认值**：使用 `.get(key, default)`
4. **调试优先**：不确定格式时，先运行调试代码

---

## 📚 相关文件

- `Dify代码节点_处理HTTP响应_修复KeyError.py` - 修复后的代码
- `dify_local_service_flexible.py` - 查看实际响应格式

---

## 🎯 总结

**问题**：代码访问了不存在的键 `data['data']`

**解决**：
1. ✅ 使用 `.get()` 方法访问键
2. ✅ 根据实际响应格式修改代码
3. ✅ 添加类型检查和错误处理

**关键**：根据实际服务返回的格式编写代码，不要假设格式！

