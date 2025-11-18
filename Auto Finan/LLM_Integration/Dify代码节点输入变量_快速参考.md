# Dify 代码节点输入变量快速参考

## 🚀 3 步配置

### 步骤 1：定义变量（在开始节点或上游节点）

**开始节点**：
- 添加输入变量：`excel_path`, `sheet_name`, `serial`

**或上游节点输出**：
- HTTP 请求节点自动输出：`http_response`

---

### 步骤 2：连接节点

- 将上游节点连接到代码节点
- 变量会自动传递

---

### 步骤 3：在代码中访问

```python
# 基本用法
variable = inputs.get('variable_name', 'default_value')

# 多个变量
var1 = inputs.get('var1', '')
var2 = inputs.get('var2', '')
var3 = inputs.get('var3', '')

# 验证
if not all([var1, var2, var3]):
    output = {"error": "缺少变量"}
else:
    output = {"result": "成功"}
```

---

## 📝 常用模式

### 模式 1：从开始节点获取

```python
excel_path = inputs.get('excel_path', '')
sheet_name = inputs.get('sheet_name', '')
serial = inputs.get('serial', '')
```

### 模式 2：从 HTTP 响应获取

```python
import json

http_response = inputs.get('http_response', '')
response = json.loads(http_response) if isinstance(http_response, str) else http_response
```

### 模式 3：安全获取（兼容不同版本）

```python
def get_inputs_safe():
    try:
        if 'inputs' in globals():
            return globals()['inputs']
    except:
        pass
    return {}

inputs = get_inputs_safe()
variable = inputs.get('variable_name', '')
```

---

## ⚠️ 重要提示

1. **代码节点**：使用 `inputs.get('variable_name')`
2. **HTTP 请求节点**：使用 `{{#workflow.variable_name#}}`
3. **变量名区分大小写**
4. **总是提供默认值**

---

## 🔍 调试代码

```python
output = {
    "debug": {
        "inputs": inputs if 'inputs' in globals() else "未定义",
        "keys": list(inputs.keys()) if 'inputs' in globals() else []
    }
}
```

