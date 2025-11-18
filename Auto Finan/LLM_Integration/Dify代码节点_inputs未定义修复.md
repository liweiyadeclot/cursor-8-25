# Dify 代码节点 inputs 未定义错误修复

## ❌ 错误信息

```
NameError: name 'inputs' is not defined
```

---

## 🔍 问题原因

不同版本的 Dify 可能使用不同的变量访问方式。`inputs` 可能不是自动可用的。

---

## ✅ 解决方案

### 方案 1：使用全局变量（推荐）

在代码开头添加变量定义，兼容不同版本：

```python
import json

# 兼容不同 Dify 版本的变量访问方式
try:
    # 方式 1：直接使用 inputs（某些版本）
    if 'inputs' in globals():
        pass  # inputs 已存在
    else:
        # 方式 2：从上下文获取（某些版本）
        try:
            from dify_workflow import get_inputs
            inputs = get_inputs()
        except:
            # 方式 3：使用 locals() 或 globals()
            inputs = locals().get('inputs', {})
            if not inputs:
                inputs = globals().get('inputs', {})
except:
    inputs = {}

# 获取 HTTP 响应
http_response = inputs.get('http_response', '')

# 解析 JSON
if isinstance(http_response, str):
    try:
        response = json.loads(http_response)
    except json.JSONDecodeError as e:
        output = {
            "success": False,
            "error": f"JSON 解析失败: {str(e)}",
            "raw_response": http_response[:200]
        }
else:
    response = http_response

# 处理响应
if not response:
    output = {
        "success": False,
        "error": "响应为空"
    }
elif response.get("success"):
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

### 方案 2：使用函数参数（如果支持）

某些 Dify 版本可能通过函数参数传递变量：

```python
import json

def process_response(inputs=None):
    if inputs is None:
        # 尝试从全局获取
        inputs = globals().get('inputs', {})
    
    http_response = inputs.get('http_response', '')
    
    # 解析和处理...
    if isinstance(http_response, str):
        response = json.loads(http_response)
    else:
        response = http_response
    
    if response.get("success"):
        return {
            "success": True,
            "mcp_prompt": response.get("mcp_prompt", "")
        }
    else:
        return {
            "success": False,
            "error": response.get("error", "未知错误")
        }

# 调用函数
output = process_response()
```

---

### 方案 3：直接从 HTTP 响应节点获取（最简单）

如果 HTTP 请求节点直接输出变量，可以直接使用：

```python
import json

# 尝试多种方式获取响应
http_response = None

# 方式 1：从 inputs
try:
    http_response = inputs.get('http_response', '')
except NameError:
    pass

# 方式 2：从全局变量
if not http_response:
    http_response = globals().get('http_response', '')

# 方式 3：从 HTTP 请求节点的输出变量名
# 检查 Dify 中 HTTP 请求节点的输出变量名
# 可能是：response, result, data 等
if not http_response:
    for var_name in ['http_response', 'response', 'result', 'data']:
        try:
            http_response = globals().get(var_name, '')
            if http_response:
                break
        except:
            pass

# 如果还是找不到，尝试解析所有可用变量
if not http_response:
    # 输出调试信息
    output = {
        "success": False,
        "error": "无法找到 HTTP 响应",
        "debug": {
            "globals_keys": list(globals().keys()),
            "locals_keys": list(locals().keys())
        }
    }
else:
    # 解析 JSON
    if isinstance(http_response, str):
        try:
            response = json.loads(http_response)
        except:
            output = {
                "success": False,
                "error": "JSON 解析失败",
                "raw": str(http_response)[:200]
            }
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

---

### 方案 4：检查 Dify 版本和配置

在代码节点中添加调试代码，查看可用变量：

```python
# 调试：查看所有可用变量
debug_info = {
    "globals": list(globals().keys()),
    "locals": list(locals().keys()),
    "dir": dir()
}

output = {
    "debug": debug_info
}
```

运行后查看输出，找到正确的变量名。

---

## 🎯 推荐方案（最兼容）

使用以下代码，兼容大多数 Dify 版本：

```python
import json

# 安全获取 inputs
def get_inputs_safe():
    """安全获取 inputs，兼容不同 Dify 版本"""
    # 方式 1：直接使用
    try:
        if 'inputs' in globals():
            return globals()['inputs']
    except:
        pass
    
    # 方式 2：从 locals
    try:
        if 'inputs' in locals():
            return locals()['inputs']
    except:
        pass
    
    # 方式 3：尝试导入
    try:
        import dify_workflow
        return dify_workflow.get_inputs()
    except:
        pass
    
    # 方式 4：返回空字典
    return {}

# 获取 inputs
inputs = get_inputs_safe()

# 获取 HTTP 响应
http_response = inputs.get('http_response', '')

# 如果为空，尝试其他可能的变量名
if not http_response:
    for key in ['http_response', 'response', 'result', 'data']:
        if key in inputs:
            http_response = inputs[key]
            break

# 解析 JSON
if isinstance(http_response, str):
    try:
        response = json.loads(http_response)
    except json.JSONDecodeError as e:
        output = {
            "success": False,
            "error": f"JSON 解析失败: {str(e)}",
            "raw_response": str(http_response)[:200]
        }
else:
    response = http_response if http_response else {}

# 处理响应
if response.get("success"):
    output = {
        "success": True,
        "mcp_prompt": response.get("mcp_prompt", ""),
        "prompt_length": response.get("prompt_length", 0)
    }
else:
    output = {
        "success": False,
        "error": response.get("error", "未知错误"),
        "debug": {
            "response_keys": list(response.keys()) if isinstance(response, dict) else [],
            "inputs_keys": list(inputs.keys())
        }
    }
```

---

## 🔍 调试步骤

1. **添加调试代码**：
   ```python
   output = {
       "debug": {
           "globals": list(globals().keys()),
           "has_inputs": 'inputs' in globals()
       }
   }
   ```

2. **查看输出**，找到正确的变量名

3. **根据输出调整代码**

---

## 📝 完整代码（直接使用）

```python
import json

# 安全获取变量
try:
    inputs = globals().get('inputs', locals().get('inputs', {}))
except:
    inputs = {}

# 获取 HTTP 响应
http_response = inputs.get('http_response', inputs.get('response', ''))

# 解析 JSON
if isinstance(http_response, str):
    try:
        response = json.loads(http_response)
    except:
        output = {
            "success": False,
            "error": "JSON 解析失败",
            "raw": str(http_response)[:200]
        }
else:
    response = http_response if http_response else {}

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

---

## ✅ 如果仍然失败

请提供：
1. Dify 版本信息
2. 代码节点的完整错误信息
3. 运行调试代码后的输出

这样我可以提供更精确的解决方案。

