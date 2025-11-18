# Dify "Output error is missing" 错误修复

## ❌ 错误信息

```
Output error is missing
```

---

## 🔍 问题原因

在 Dify 代码节点中，如果：
1. 函数没有返回值
2. 返回的字典中缺少必需的输出变量
3. 输出变量配置不正确

就会报 "Output error is missing" 错误。

---

## ✅ 解决方案

### 方案 1：确保所有输出变量都存在

在 `return` 字典中，确保所有在 Dify 中配置的输出变量都存在：

```python
def main(http_response: str) -> dict:
    import json
    
    try:
        response = json.loads(http_response)
    except json.JSONDecodeError as e:
        return {
            "success": False,
            "error": f"JSON 解析失败: {str(e)}",
            "mcp_prompt": "",  # 确保所有输出变量都存在
            "prompt_length": 0
        }
    
    if response.get("success"):
        return {
            "success": True,
            "mcp_prompt": response.get("mcp_prompt", ""),
            "prompt_length": response.get("prompt_length", 0)
        }
    else:
        return {
            "success": False,
            "error": response.get("error", "未知错误"),
            "mcp_prompt": "",  # 确保所有输出变量都存在
            "prompt_length": 0
        }
```

---

### 方案 2：检查输出变量配置

在 Dify 代码节点的配置中：

1. **找到 "输出变量" 或 "Output Variables" 部分**
2. **确保所有输出变量都已声明**：
   - `success` (boolean)
   - `mcp_prompt` (string)
   - `prompt_length` (number)
   - `error` (string, 可选)

3. **确保代码返回的字典包含所有这些变量**

---

### 方案 3：简化输出变量

如果不需要所有变量，可以只返回必需的：

```python
def main(http_response: str) -> dict:
    import json
    
    try:
        response = json.loads(http_response)
    except:
        return {
            "success": False,
            "mcp_prompt": ""  # 只返回必需的变量
        }
    
    if response.get("success"):
        return {
            "success": True,
            "mcp_prompt": response.get("mcp_prompt", "")
        }
    else:
        return {
            "success": False,
            "mcp_prompt": ""
        }
```

然后在 Dify 中只配置这两个输出变量。

---

## 📝 完整代码（推荐）

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
            "mcp_prompt": "",
            "prompt_length": 0
        }
    
    # 检查响应格式
    if not isinstance(response, dict):
        return {
            "success": False,
            "error": f"响应格式错误",
            "mcp_prompt": "",
            "prompt_length": 0
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
            "mcp_prompt": "",
            "prompt_length": 0
        }
```

---

## 🔧 在 Dify 中配置输出变量

### 步骤 1：添加输出变量

在代码节点的配置中：

1. 找到 **"输出变量"** 或 **"Output Variables"** 部分
2. 添加以下变量：

| 变量名 | 类型 | 说明 |
|--------|------|------|
| `success` | boolean | 是否成功 |
| `mcp_prompt` | string | MCP 提示词 |
| `prompt_length` | number | 提示词长度 |
| `error` | string | 错误信息（可选） |

### 步骤 2：确保代码返回所有变量

代码中的 `return` 字典必须包含所有这些变量。

---

## ⚠️ 常见错误

### 错误 1：缺少输出变量

**现象**：代码返回的字典中缺少某个输出变量

**解决**：确保所有输出变量都存在，即使值为空

```python
# ❌ 错误：缺少 mcp_prompt
return {
    "success": False,
    "error": "错误信息"
}

# ✅ 正确：包含所有输出变量
return {
    "success": False,
    "error": "错误信息",
    "mcp_prompt": "",  # 确保存在
    "prompt_length": 0
}
```

---

### 错误 2：输出变量类型不匹配

**现象**：代码返回的类型与配置的类型不一致

**解决**：确保类型匹配

```python
# 确保类型正确
return {
    "success": True,  # boolean
    "mcp_prompt": "",  # string
    "prompt_length": 0  # number
}
```

---

### 错误 3：函数没有返回值

**现象**：函数在某些情况下没有 return

**解决**：确保所有代码路径都有 return

```python
# ❌ 错误：某些路径没有 return
if condition:
    return {"result": "ok"}
# 缺少 else 分支的 return

# ✅ 正确：所有路径都有 return
if condition:
    return {"result": "ok"}
else:
    return {"result": "fail"}
```

---

## 🎯 调试方法

### 方法 1：简化代码测试

先使用最简单的代码测试：

```python
def main(http_response: str) -> dict:
    return {
        "success": True,
        "mcp_prompt": "test",
        "prompt_length": 4
    }
```

如果这个能工作，说明输出变量配置正确，问题在代码逻辑。

---

### 方法 2：查看实际响应

```python
def main(http_response: str) -> dict:
    import json
    
    # 返回原始响应用于调试
    try:
        response = json.loads(http_response)
        return {
            "debug": response,  # 查看实际响应
            "success": True,
            "mcp_prompt": "",
            "prompt_length": 0
        }
    except:
        return {
            "debug": http_response[:200],
            "success": False,
            "mcp_prompt": "",
            "prompt_length": 0
        }
```

---

## ✅ 检查清单

- [ ] 函数名是 `main`
- [ ] 函数有返回值（所有路径都有 return）
- [ ] 返回的字典包含所有输出变量
- [ ] 输出变量类型正确
- [ ] 在 Dify 中配置了所有输出变量
- [ ] 输出变量名与代码中的键名匹配

---

## 📚 相关文件

- `Dify代码节点_处理HTTP响应_最终版.py` - 修复后的代码
- `Dify代码节点_处理HTTP响应_官方标准版.py` - 官方标准版

---

## 🎉 总结

**问题**：代码节点没有正确返回输出变量

**解决**：
1. ✅ 确保所有输出变量都存在
2. ✅ 检查输出变量配置
3. ✅ 确保所有代码路径都有 return

**关键**：返回的字典必须包含所有在 Dify 中配置的输出变量！

