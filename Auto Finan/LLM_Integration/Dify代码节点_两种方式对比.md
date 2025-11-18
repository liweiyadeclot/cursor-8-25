# Dify 代码节点两种输出方式对比

## 🔍 问题

"Output error is missing" 错误可能是因为：
1. 使用了 `return`，但 Dify 期望 `output` 变量
2. 或者相反，使用了 `output`，但 Dify 期望 `return`

---

## 📊 两种方式对比

### 方式 1：使用 `output` 变量（某些版本）

```python
import json

# 获取输入
http_response = http_response  # 参数会自动传入

# 处理
response = json.loads(http_response)

# 设置 output
if response.get("success"):
    output = {
        "success": True,
        "mcp_prompt": response.get("mcp_prompt", ""),
        "prompt_length": response.get("prompt_length", 0)
    }
else:
    output = {
        "success": False,
        "mcp_prompt": "",
        "prompt_length": 0
    }
```

**特点**：
- 不使用函数定义
- 直接使用 `output` 变量
- 代码在模块级别执行

---

### 方式 2：使用 `return`（标准方式）

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

**特点**：
- 使用函数定义
- 使用 `return` 返回
- 符合官方文档

---

## 🎯 推荐测试顺序

### 测试 1：使用 output 变量（先试这个）

```python
import json

response = json.loads(http_response)

output = {
    "success": response.get("success", False),
    "mcp_prompt": response.get("mcp_prompt", ""),
    "prompt_length": response.get("prompt_length", 0)
}
```

---

### 测试 2：使用 return（如果测试 1 失败）

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

## 🔧 兼容两种方式的代码

```python
import json

# 获取输入
try:
    # 尝试从函数参数获取
    if 'http_response' in locals():
        http_response = locals()['http_response']
    elif 'http_response' in globals():
        http_response = globals()['http_response']
except:
    http_response = None

# 处理
if http_response:
    try:
        response = json.loads(http_response)
        
        result = {
            "success": response.get("success", False),
            "mcp_prompt": response.get("mcp_prompt", ""),
            "prompt_length": response.get("prompt_length", 0)
        }
        
        # 方式 1：使用 output 变量
        output = result
        
        # 方式 2：如果使用函数，也可以 return
        # return result
        
    except Exception as e:
        output = {
            "success": False,
            "mcp_prompt": "",
            "prompt_length": 0,
            "error": str(e)
        }
else:
    output = {
        "success": False,
        "mcp_prompt": "",
        "prompt_length": 0,
        "error": "无法获取 http_response"
    }
```

---

## 📝 最简单的测试代码

### 测试 output 方式

```python
output = {
    "success": True,
    "mcp_prompt": "test",
    "prompt_length": 4
}
```

### 测试 return 方式

```python
def main(http_response: str) -> dict:
    return {
        "success": True,
        "mcp_prompt": "test",
        "prompt_length": 4
    }
```

---

## ✅ 推荐方案

根据你的 Dify 版本，尝试以下顺序：

1. **先试 output 变量方式**（最简单）
2. **如果失败，试 return 方式**
3. **如果都失败，使用兼容版本**

---

## 📚 相关文件

- `Dify代码节点_处理HTTP响应_纯output版.py` - 纯 output 版本
- `Dify代码节点_处理HTTP响应_官方标准版.py` - return 版本

