# Dify "Output error is missing" 详细排查

## ❌ 错误信息

```
Output error is missing
```

---

## 🔍 可能的原因

1. **输出变量未在 Dify 中配置**
2. **代码返回的字典缺少必需的输出变量**
3. **输出变量名不匹配**
4. **函数执行出错，没有返回值**

---

## ✅ 排查步骤

### 步骤 1：检查输出变量配置

在 Dify 代码节点中：

1. **找到 "输出变量" 或 "Output Variables" 部分**
2. **检查是否配置了以下变量**：
   - `success` (boolean)
   - `mcp_prompt` (string)
   - `prompt_length` (number)

3. **如果没有配置，添加这些变量**

---

### 步骤 2：使用调试版本

先使用调试版本查看实际输入：

```python
def main(http_response: str) -> dict:
    import json
    
    try:
        response = json.loads(http_response)
        
        # 返回调试信息
        return {
            "debug_response_keys": list(response.keys()) if isinstance(response, dict) else [],
            "debug_success": response.get("success") if isinstance(response, dict) else None,
            "success": True,
            "mcp_prompt": response.get("mcp_prompt", "") if isinstance(response, dict) else "",
            "prompt_length": response.get("prompt_length", 0) if isinstance(response, dict) else 0
        }
    except Exception as e:
        return {
            "debug_error": str(e),
            "success": False,
            "mcp_prompt": "",
            "prompt_length": 0,
            "error": f"处理失败: {str(e)}"
        }
```

运行后查看输出，确认：
- 输入是否正确
- 响应格式是否正确
- 是否有错误

---

### 步骤 3：使用最简版本

如果调试版本能工作，使用最简版本：

```python
def main(http_response: str) -> dict:
    import json
    
    response = json.loads(http_response)
    
    if response.get("success"):
        return {
            "success": True,
            "mcp_prompt": response.get("mcp_prompt", ""),
            "prompt_length": response.get("prompt_length", 0)
        }
    else:
        return {
            "success": False,
            "mcp_prompt": "",
            "prompt_length": 0
        }
```

---

### 步骤 4：检查 Dify 版本差异

不同版本的 Dify 可能有不同的要求：

#### 方式 A：使用 output 变量（某些版本）

```python
import json

response = json.loads(http_response)

output = {
    "success": response.get("success", False),
    "mcp_prompt": response.get("mcp_prompt", ""),
    "prompt_length": response.get("prompt_length", 0)
}
```

#### 方式 B：使用 return（标准方式）

```python
def main(http_response: str) -> dict:
    import json
    response = json.loads(http_response)
    return {
        "success": response.get("success", False),
        "mcp_prompt": response.get("mcp_prompt", ""),
        "prompt_length": response.get("prompt_length", 0)
    }
```

---

## 🔧 完整解决方案

### 方案 1：确保输出变量配置正确

在 Dify 代码节点配置中：

1. **输入变量**：
   - `http_response` (string)

2. **输出变量**（必须配置）：
   - `success` (boolean)
   - `mcp_prompt` (string)
   - `prompt_length` (number)

3. **代码**：
```python
def main(http_response: str) -> dict:
    import json
    
    response = json.loads(http_response)
    
    return {
        "success": response.get("success", False),
        "mcp_prompt": response.get("mcp_prompt", ""),
        "prompt_length": response.get("prompt_length", 0)
    }
```

---

### 方案 2：兼容不同 Dify 版本

```python
import json

# 获取输入
http_response = http_response  # 参数会自动传入

# 解析
response = json.loads(http_response)

# 设置输出（某些版本需要）
output = {
    "success": response.get("success", False),
    "mcp_prompt": response.get("mcp_prompt", ""),
    "prompt_length": response.get("prompt_length", 0)
}

# 如果使用函数方式
def main(http_response: str) -> dict:
    import json
    response = json.loads(http_response)
    return output
```

---

## 📝 测试代码

### 测试 1：最简单的版本

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

### 测试 2：检查输入

```python
def main(http_response: str) -> dict:
    return {
        "debug": http_response[:100],  # 查看前100字符
        "success": True,
        "mcp_prompt": "",
        "prompt_length": 0
    }
```

---

## ⚠️ 常见问题

### 问题 1：输出变量未配置

**现象**：代码返回了字典，但 Dify 报错

**解决**：在 Dify 中配置输出变量

---

### 问题 2：输出变量名不匹配

**现象**：代码返回 `mcp_prompt`，但 Dify 配置的是 `prompt`

**解决**：确保变量名完全匹配

---

### 问题 3：函数执行出错

**现象**：代码有语法错误或运行时错误

**解决**：
1. 检查代码语法
2. 添加 try-except 处理错误
3. 确保所有路径都有返回值

---

## 🎯 推荐方案

### 最保险的方式

```python
def main(http_response: str) -> dict:
    import json
    
    try:
        response = json.loads(http_response)
        
        # 确保所有输出变量都存在
        result = {
            "success": response.get("success", False),
            "mcp_prompt": response.get("mcp_prompt", ""),
            "prompt_length": response.get("prompt_length", 0)
        }
        
        # 如果有错误信息，也添加
        if not result["success"]:
            result["error"] = response.get("error", "未知错误")
        
        return result
        
    except Exception as e:
        # 错误时也要返回所有输出变量
        return {
            "success": False,
            "mcp_prompt": "",
            "prompt_length": 0,
            "error": str(e)
        }
```

---

## 📚 相关文件

- `Dify代码节点_处理HTTP响应_简化版.py` - 最简版本
- `Dify代码节点_调试版.py` - 调试版本

---

## 💡 如果仍然失败

请提供：
1. Dify 版本信息
2. 代码节点的完整配置截图（包括输入变量和输出变量）
3. 运行调试版本后的输出

这样我可以提供更精确的解决方案。

