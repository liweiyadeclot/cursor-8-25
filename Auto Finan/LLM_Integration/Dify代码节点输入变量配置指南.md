# Dify 代码节点输入变量配置指南

## 📋 概述

在 Dify 中，代码节点的输入变量有两种来源：
1. **从开始节点传入**（工作流输入）
2. **从上游节点传入**（节点输出）

---

## 🎯 方法 1：从开始节点获取变量

### 步骤 1：在开始节点定义输入变量

1. 添加 **开始节点**
2. 点击节点，在右侧面板中找到 **"输入变量"** 或 **"Variables"** 部分
3. 点击 **"添加变量"** 或 **"+"** 按钮
4. 添加以下变量：

| 变量名 | 类型 | 说明 | 示例值 |
|--------|------|------|--------|
| `excel_path` | 字符串 | Excel 文件路径 | `C:\Users\FH\source\repos\Auto Finan\Auto Finan\420财务050823.xlsx` |
| `sheet_name` | 字符串 | 工作表名称 | `3-报销` |
| `serial` | 字符串 | 序号 | `1` |

### 步骤 2：在代码节点中访问变量

**代码节点配置**：

```python
# 从开始节点获取变量
excel_path = inputs.get('excel_path', '')
sheet_name = inputs.get('sheet_name', '')
serial = inputs.get('serial', '')

# 验证变量
if not excel_path or not sheet_name or not serial:
    output = {
        "success": False,
        "error": "缺少必要参数"
    }
else:
    # 处理逻辑
    output = {
        "success": True,
        "excel_path": excel_path,
        "sheet_name": sheet_name,
        "serial": serial
    }
```

---

## 🎯 方法 2：从上游节点获取变量

### 场景：从 HTTP 请求节点获取响应

#### 步骤 1：HTTP 请求节点输出变量

HTTP 请求节点会自动创建输出变量：
- `http_response` - HTTP 响应内容（通常是 JSON 字符串）

#### 步骤 2：在代码节点中访问

**代码节点配置**：

```python
import json

# 从 HTTP 请求节点获取响应
http_response = inputs.get('http_response', '')

# 解析 JSON
if isinstance(http_response, str):
    try:
        response = json.loads(http_response)
    except json.JSONDecodeError as e:
        output = {
            "success": False,
            "error": f"JSON 解析失败: {str(e)}"
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

## 🔧 在 Dify 界面中配置输入变量

### 方法 A：通过节点连接自动传递

1. **连接节点**：
   - 将上游节点的输出连接到代码节点
   - Dify 会自动将上游节点的输出变量传递给代码节点

2. **在代码中访问**：
   ```python
   # 变量会自动出现在 inputs 字典中
   variable_name = inputs.get('variable_name', '')
   ```

### 方法 B：手动配置输入变量（某些版本）

1. **选择代码节点**
2. **找到 "输入变量" 或 "Input Variables" 部分**
3. **点击 "添加变量"**
4. **配置变量**：
   - **变量名**：`excel_path`
   - **来源**：选择上游节点
   - **字段**：选择要获取的字段名

---

## 📝 完整示例

### 示例 1：从开始节点获取变量

**工作流结构**：
```
开始节点 → 代码节点
```

**开始节点配置**：
- 输入变量：`excel_path`, `sheet_name`, `serial`

**代码节点代码**：
```python
# 获取输入变量
excel_path = inputs.get('excel_path', '')
sheet_name = inputs.get('sheet_name', '')
serial = inputs.get('serial', '')

# 验证
if not all([excel_path, sheet_name, serial]):
    output = {
        "success": False,
        "error": "缺少必要参数",
        "received": {
            "excel_path": excel_path,
            "sheet_name": sheet_name,
            "serial": serial
        }
    }
else:
    # 处理逻辑
    output = {
        "success": True,
        "excel_path": excel_path,
        "sheet_name": sheet_name,
        "serial": serial
    }
```

---

### 示例 2：从 HTTP 请求节点获取响应

**工作流结构**：
```
开始节点 → HTTP请求节点 → 代码节点
```

**HTTP 请求节点**：
- 输出变量：`http_response`

**代码节点代码**：
```python
import json

# 安全获取 inputs
def get_inputs_safe():
    try:
        if 'inputs' in globals():
            return globals()['inputs']
    except:
        pass
    try:
        if 'inputs' in locals():
            return locals()['inputs']
    except:
        pass
    return {}

inputs = get_inputs_safe()

# 获取 HTTP 响应
http_response = inputs.get('http_response', '')

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

## 🔍 调试：查看可用变量

如果不知道变量名，可以使用以下调试代码：

```python
# 调试：查看所有可用变量
output = {
    "debug": {
        "inputs_keys": list(inputs.keys()) if 'inputs' in globals() or 'inputs' in locals() else [],
        "inputs": inputs if 'inputs' in globals() or 'inputs' in locals() else "inputs 未定义",
        "globals_keys": [k for k in globals().keys() if not k.startswith('_')][:20]
    }
}
```

运行后查看输出，找到正确的变量名。

---

## ⚠️ 常见问题

### 问题 1：inputs 未定义

**错误**：`NameError: name 'inputs' is not defined`

**解决**：使用安全获取方法：

```python
def get_inputs_safe():
    try:
        if 'inputs' in globals():
            return globals()['inputs']
    except:
        pass
    return {}

inputs = get_inputs_safe()
```

---

### 问题 2：变量值为空

**原因**：
- 变量名不匹配
- 上游节点未输出该变量
- 变量未正确传递

**解决**：
1. 检查变量名是否完全匹配（区分大小写）
2. 检查上游节点是否正确输出变量
3. 使用调试代码查看可用变量

---

### 问题 3：变量类型错误

**原因**：变量类型与预期不符

**解决**：在代码中转换类型：

```python
# 转换为字符串
serial = str(inputs.get('serial', ''))

# 转换为数字
count = int(inputs.get('count', 0))

# 转换为布尔值
is_valid = bool(inputs.get('is_valid', False))
```

---

## 📊 变量访问方式对比

| 节点类型 | 变量来源 | 访问方式 | 示例 |
|---------|---------|---------|------|
| **代码节点** | 开始节点 | `inputs.get('variable_name')` | `excel_path = inputs.get('excel_path', '')` |
| **代码节点** | 上游节点 | `inputs.get('output_variable')` | `http_response = inputs.get('http_response', '')` |
| **HTTP 请求节点** | 工作流变量 | `{{#workflow.variable_name#}}` | `"path": "{{#workflow.excel_path#}}"` |
| **条件判断节点** | 工作流变量 | `{{#workflow.variable_name#}}` | `{{#workflow.success#}} == true` |

---

## ✅ 配置检查清单

- [ ] 开始节点已定义输入变量
- [ ] 代码节点已连接到上游节点（如果需要）
- [ ] 代码中使用 `inputs.get()` 访问变量
- [ ] 变量名完全匹配（区分大小写）
- [ ] 提供了默认值（使用 `.get()` 方法）
- [ ] 添加了变量验证逻辑

---

## 🎯 快速参考

### 基本模式

```python
# 1. 获取变量
variable = inputs.get('variable_name', 'default_value')

# 2. 验证变量
if not variable:
    output = {"error": "变量为空"}
else:
    # 3. 处理逻辑
    output = {"result": variable}
```

### 多个变量

```python
# 获取多个变量
var1 = inputs.get('var1', '')
var2 = inputs.get('var2', '')
var3 = inputs.get('var3', '')

# 验证
if not all([var1, var2, var3]):
    output = {"error": "缺少变量"}
else:
    output = {"result": f"{var1}-{var2}-{var3}"}
```

---

## 📚 相关文档

- `Dify代码节点正确语法.md` - 变量访问语法
- `Dify工作流完整配置指南.md` - 完整工作流配置
- `Dify代码节点_inputs未定义修复.md` - inputs 未定义问题修复

---

## 💡 最佳实践

1. **总是使用 `.get()` 方法**：提供默认值，避免 KeyError
2. **验证变量**：在使用前检查变量是否存在和有效
3. **使用调试代码**：如果不确定变量名，先运行调试代码
4. **统一命名**：保持变量名一致，避免大小写问题
5. **处理类型**：明确变量的类型，必要时进行转换

---

## 🎉 总结

在 Dify 代码节点中配置输入变量：

1. **从开始节点**：在开始节点定义变量，在代码中使用 `inputs.get('variable_name')`
2. **从上游节点**：连接上游节点，使用 `inputs.get('output_variable')`
3. **安全访问**：使用 `inputs.get()` 并提供默认值
4. **调试**：使用调试代码查看可用变量

记住：**代码节点使用 `inputs` 字典，HTTP 请求节点使用 `{{#workflow.variable_name#}}` 模板语法**。

