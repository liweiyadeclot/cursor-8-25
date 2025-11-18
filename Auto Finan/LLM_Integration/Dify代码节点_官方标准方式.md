# Dify 代码节点官方标准方式

## 📚 根据 Dify 官方文档

Dify 代码节点应该使用 **函数定义** 的方式，而不是直接使用 `inputs` 字典。

---

## ✅ 官方标准格式

### 基本格式

```python
def main(参数名: 类型) -> 返回类型:
    # 处理逻辑
    return {
        'output_variable': value
    }
```

---

## 📝 完整示例

### 示例 1：处理 HTTP 响应

**官方文档示例**：
```python
def main(http_response: str) -> dict:
    import json
    
    data = json.loads(http_response)
    
    return {
        # 注意在输出变量中声明 result
        'result': data['data']['name'] 
    }
```

**我们的应用场景**：
```python
def main(http_response: str) -> dict:
    import json
    
    # 解析 JSON
    try:
        response = json.loads(http_response)
    except json.JSONDecodeError as e:
        return {
            "success": False,
            "error": f"JSON 解析失败: {str(e)}"
        }
    
    # 处理响应
    if response.get("success"):
        return {
            "success": True,
            "mcp_prompt": response.get("mcp_prompt", ""),
            "prompt_length": response.get("prompt_length", 0)
        }
    else:
        return {
            "success": False,
            "error": response.get("error", "未知错误")
        }
```

---

### 示例 2：从开始节点获取多个变量

```python
def main(excel_path: str, sheet_name: str, serial: str) -> dict:
    """
    处理 Excel 数据
    
    参数：
        excel_path: Excel 文件路径
        sheet_name: 工作表名称
        serial: 序号
    """
    # 验证参数
    if not all([excel_path, sheet_name, serial]):
        return {
            "success": False,
            "error": "缺少必要参数"
        }
    
    # 处理逻辑
    # ...
    
    return {
        "success": True,
        "result": "处理完成"
    }
```

---

## 🔧 在 Dify 中配置

### 步骤 1：定义函数参数

在代码节点中，函数参数名对应输入变量名：

```python
def main(http_response: str) -> dict:
    # http_response 对应输入变量名
    pass
```

### 步骤 2：配置输入变量

在 Dify 代码节点的配置中：

1. **输入变量**：
   - 变量名：`http_response`
   - 类型：`string`
   - 来源：选择上游节点（如 HTTP 请求节点）

2. **输出变量**：
   - 在 `return` 字典中定义的键会自动成为输出变量
   - 例如：`return {"success": True, "mcp_prompt": "..."}`
   - 输出变量：`success`, `mcp_prompt`

---

## 📊 对比：旧方式 vs 官方方式

### 旧方式（不推荐）

```python
# ❌ 直接使用 inputs（可能不兼容）
http_response = inputs.get('http_response', '')
response = json.loads(http_response)
output = {"result": response}
```

### 官方方式（推荐）

```python
# ✅ 使用函数定义
def main(http_response: str) -> dict:
    import json
    response = json.loads(http_response)
    return {"result": response}
```

---

## 🎯 完整工作流示例

### 节点 1：开始节点

**输入变量**：
- `excel_path` (string)
- `sheet_name` (string)
- `serial` (string)

---

### 节点 2：HTTP 请求节点

**配置**：
- URL: `http://192.168.137.133:8001/api/excel-to-prompt`
- 请求体: 
```json
{
  "excel_path": "{{#workflow.excel_path#}}",
  "sheet_name": "{{#workflow.sheet_name#}}",
  "serial": "{{#workflow.serial#}}"
}
```

**输出变量**：
- `http_response` (string) - 自动创建

---

### 节点 3：代码节点（处理响应）

**输入变量配置**：
- `http_response` (string) - 来自 HTTP 请求节点

**代码**：
```python
def main(http_response: str) -> dict:
    import json
    
    # 解析 JSON
    try:
        response = json.loads(http_response)
    except json.JSONDecodeError as e:
        return {
            "success": False,
            "error": f"JSON 解析失败: {str(e)}"
        }
    
    # 处理响应
    if response.get("success"):
        return {
            "success": True,
            "mcp_prompt": response.get("mcp_prompt", ""),
            "prompt_length": response.get("prompt_length", 0)
        }
    else:
        return {
            "success": False,
            "error": response.get("error", "未知错误")
        }
```

**输出变量**（自动从 return 字典创建）：
- `success` (boolean)
- `mcp_prompt` (string)
- `prompt_length` (number)
- `error` (string, 如果失败)

---

## ⚠️ 重要提示

1. **函数名必须是 `main`**
2. **参数名必须与输入变量名匹配**
3. **返回类型使用 `-> dict`**
4. **在输出变量中声明返回的键**（某些版本需要）

---

## 🔍 多参数示例

```python
def main(excel_path: str, sheet_name: str, serial: str) -> dict:
    """
    处理 Excel 数据
    
    参数对应输入变量：
    - excel_path: 输入变量名
    - sheet_name: 输入变量名
    - serial: 输入变量名
    """
    # 验证
    if not all([excel_path, sheet_name, serial]):
        return {
            "success": False,
            "error": "缺少必要参数"
        }
    
    # 处理逻辑
    # ...
    
    return {
        "success": True,
        "result": "处理完成"
    }
```

**输入变量配置**：
- `excel_path` (string) - 来自开始节点
- `sheet_name` (string) - 来自开始节点
- `serial` (string) - 来自开始节点

---

## 📝 类型提示

支持的类型：
- `str` - 字符串
- `int` - 整数
- `float` - 浮点数
- `bool` - 布尔值
- `dict` - 字典
- `list` - 列表

示例：
```python
def main(
    name: str,
    age: int,
    score: float,
    is_active: bool,
    data: dict,
    items: list
) -> dict:
    return {"result": "success"}
```

---

## ✅ 配置检查清单

- [ ] 使用 `def main()` 函数定义
- [ ] 参数名与输入变量名匹配
- [ ] 添加类型提示
- [ ] 返回类型使用 `-> dict`
- [ ] 在 Dify 中配置了输入变量
- [ ] 输入变量名与函数参数名一致

---

## 🎉 总结

**官方标准方式**：
1. 使用 `def main(参数: 类型) -> dict:` 定义函数
2. 参数名对应输入变量名
3. 返回字典，键自动成为输出变量
4. 在 Dify 中配置输入变量，名称与函数参数匹配

**优势**：
- ✅ 符合官方规范
- ✅ 类型安全
- ✅ 更清晰
- ✅ 更好的 IDE 支持

---

## 📚 相关文件

- `Dify代码节点_处理HTTP响应_官方标准版.py` - 标准版代码
- `Dify代码节点_处理HTTP响应_修复版.py` - 兼容版代码（如果官方方式不工作）

