# Dify 400 错误详细排查

## ❌ 错误：Request failed with status code 400

根据诊断，你的 Excel 文件和数据都是正常的。400 错误可能是 Dify 传递的参数有问题。

---

## 🔍 排查步骤

### 步骤 1：查看完整的错误响应

在 Dify 中，查看 HTTP 请求节点的**完整响应**，应该包含：

```json
{
  "success": false,
  "error": "具体错误信息",
  "received": {
    "excel_path": "...",
    "sheet_name": "...",
    "serial": "..."
  },
  "debug": {
    "file_exists": true/false,
    "file_size": 12345
  }
}
```

**关键信息**：
- `error` - 具体错误原因
- `received` - Dify 实际发送的参数
- `debug` - 调试信息

---

### 步骤 2：检查 Dify 传递的参数

在代码节点中添加调试输出：

```python
# 在发送 HTTP 请求前的代码节点中
excel_path = inputs.get('excel_path', '')
sheet_name = inputs.get('sheet_name', '')
serial = inputs.get('serial', '')

# 输出调试信息
output = {
    "debug_before_request": {
        "excel_path": excel_path,
        "excel_path_type": type(excel_path).__name__,
        "excel_path_length": len(str(excel_path)),
        "sheet_name": sheet_name,
        "sheet_name_type": type(sheet_name).__name__,
        "serial": serial,
        "serial_type": type(serial).__name__
    },
    "ready": True
}
```

查看输出，确认：
- 变量值是否正确
- 变量类型是否正确
- 是否有空值或特殊字符

---

### 步骤 3：检查 HTTP 请求节点配置

**请求体配置**：

```json
{
  "excel_path": "{{#workflow.excel_path#}}",
  "sheet_name": "{{#workflow.sheet_name#}}",
  "serial": "{{#workflow.serial#}}"
}
```

**常见问题**：

1. **变量未定义**：
   - 如果变量不存在，Dify 会传递字符串 `"{{#workflow.excel_path#}}"`
   - 检查上游节点是否正确输出变量

2. **路径格式问题**：
   - Windows 路径中的反斜杠需要转义
   - 确保路径是完整路径

3. **变量值为空**：
   - 检查变量是否为空字符串
   - 添加默认值或验证

---

### 步骤 4：使用测试脚本验证

运行诊断工具：

```bash
python diagnose_400_error.py
```

如果诊断通过，说明数据正常，问题在 Dify 配置。

---

## 🔧 解决方案

### 方案 1：在代码节点中验证和格式化参数

在发送 HTTP 请求前，先验证和格式化参数：

```python
import os

# 获取变量
excel_path = inputs.get('excel_path', '').strip()
sheet_name = inputs.get('sheet_name', '').strip()
serial = str(inputs.get('serial', '')).strip()

# 验证
errors = []
if not excel_path:
    errors.append("excel_path 为空")
if not sheet_name:
    errors.append("sheet_name 为空")
if not serial:
    errors.append("serial 为空")

if errors:
    output = {
        "success": False,
        "error": "参数验证失败",
        "errors": errors
    }
elif not os.path.exists(excel_path):
    output = {
        "success": False,
        "error": f"文件不存在: {excel_path}"
    }
else:
    # 参数正确，准备发送请求
    output = {
        "success": True,
        "excel_path": excel_path,
        "sheet_name": sheet_name,
        "serial": serial,
        "ready_for_request": True
    }
```

---

### 方案 2：查看服务日志

服务启动后会显示请求日志，查看：
- 接收到的请求体
- 处理过程中的错误
- 返回的响应

---

### 方案 3：使用 Postman 测试

使用 Postman 直接测试服务，确认服务正常：

**请求配置**：
- URL: `http://localhost:8001/api/excel-to-prompt`
- Method: `POST`
- Body (JSON):
```json
{
  "excel_path": "C:\\Users\\FH\\source\\repos\\Auto Finan\\Auto Finan\\420财务050823.xlsx",
  "sheet_name": "3-报销",
  "serial": "1"
}
```

如果 Postman 测试成功，说明服务正常，问题在 Dify 配置。

---

## 📝 正确的 Dify 工作流配置

### 节点 1：开始节点

**输入变量**：
- `excel_path` (字符串)
- `sheet_name` (字符串)
- `serial` (字符串)

---

### 节点 2：代码节点（验证参数）

```python
import os

excel_path = inputs.get('excel_path', '').strip()
sheet_name = inputs.get('sheet_name', '').strip()
serial = str(inputs.get('serial', '')).strip()

# 验证
if not excel_path or not sheet_name or not serial:
    output = {
        "success": False,
        "error": "缺少必要参数",
        "received": {
            "excel_path": excel_path,
            "sheet_name": sheet_name,
            "serial": serial
        }
    }
elif not os.path.exists(excel_path):
    output = {
        "success": False,
        "error": f"文件不存在: {excel_path}"
    }
else:
    # 参数验证通过
    output = {
        "success": True,
        "excel_path": excel_path,
        "sheet_name": sheet_name,
        "serial": serial
    }
```

---

### 节点 3：HTTP 请求节点

**URL**: `http://192.168.137.133:8001/api/excel-to-prompt`

**请求体**：
```json
{
  "excel_path": "{{#workflow.excel_path#}}",
  "sheet_name": "{{#workflow.sheet_name#}}",
  "serial": "{{#workflow.serial#}}"
}
```

---

## 🐛 常见问题

### 问题 1：变量值为空

**现象**：Dify 传递空字符串

**解决**：在代码节点中验证并设置默认值

---

### 问题 2：路径格式错误

**现象**：路径中的反斜杠未转义

**解决**：使用双反斜杠或正斜杠

---

### 问题 3：序号类型错误

**现象**：传递数字而不是字符串

**解决**：在代码节点中转换为字符串

---

## ✅ 验证清单

- [ ] 诊断工具检查通过
- [ ] Dify 变量值正确
- [ ] HTTP 请求体格式正确
- [ ] 服务正在运行
- [ ] 可以访问服务（测试 curl）

---

## 💡 推荐做法

1. **先运行诊断**：`python diagnose_400_error.py`
2. **查看完整错误**：在 Dify 中查看 HTTP 响应
3. **添加调试输出**：在代码节点中输出变量值
4. **测试服务**：使用 Postman 或测试脚本

---

## 📚 相关文件

- `diagnose_400_error.py` - 诊断工具
- `dify_local_service_flexible.py` - 灵活版本服务（已改进）
- `test_dify_local_service.py` - 测试脚本

