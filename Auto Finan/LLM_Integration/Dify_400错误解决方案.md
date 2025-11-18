# Dify 400 错误解决方案

## ❌ 错误：Request failed with status code 400

**400 错误**表示请求已接收，但处理失败。通常是因为：
- Excel 文件不存在
- 序号不存在
- 数据格式问题

---

## 🔍 故障排查步骤

### 步骤 1：检查错误详情

在 Dify 中查看完整的错误响应，通常包含：
- `error` - 错误信息
- `suggestion` - 建议
- `received` - 接收到的参数

---

### 步骤 2：验证文件路径

**常见问题**：Excel 文件路径不正确

**检查**：
1. 文件是否存在于指定路径
2. 路径格式是否正确（Windows 使用 `\\` 或 `/`）
3. 文件是否有访问权限

**测试**：
```python
import os
excel_path = "C:\\Users\\FH\\source\\repos\\Auto Finan\\Auto Finan\\420财务050823.xlsx"
print(f"文件存在: {os.path.exists(excel_path)}")
```

---

### 步骤 3：验证序号

**常见问题**：序号不存在或格式错误

**检查**：
1. Excel 中是否存在该序号
2. 序号格式（字符串 vs 数字）
3. 工作表名称是否正确

**测试**：
```python
# 在代码节点中测试
from openpyxl import load_workbook

wb = load_workbook(excel_path, data_only=True)
ws = wb[sheet_name]

# 查找所有序号
serials = []
for row in ws.iter_rows(min_row=2, values_only=True):
    if row[0]:  # 假设序号在第一列
        serials.append(str(row[0]).strip())

print(f"可用序号: {serials}")
```

---

### 步骤 4：检查工作表名称

**常见问题**：工作表名称不匹配

**检查**：
1. 工作表名称是否正确（区分大小写）
2. 是否有空格或特殊字符

**获取所有工作表名称**：
```python
from openpyxl import load_workbook

wb = load_workbook(excel_path)
print(f"所有工作表: {wb.sheetnames}")
```

---

## ✅ 解决方案

### 方案 1：添加调试输出

在 Dify 代码节点中，添加调试信息：

```python
import os

excel_path = inputs.get('excel_path', '')
sheet_name = inputs.get('sheet_name', '')
serial = inputs.get('serial', '')

# 调试信息
debug_info = {
    "excel_path": excel_path,
    "file_exists": os.path.exists(excel_path) if excel_path else False,
    "sheet_name": sheet_name,
    "serial": serial
}

# 如果文件不存在，返回详细信息
if not os.path.exists(excel_path):
    output = {
        "success": False,
        "error": f"Excel 文件不存在: {excel_path}",
        "debug": debug_info
    }
else:
    # 继续处理...
    output = {"success": True, "debug": debug_info}
```

---

### 方案 2：验证请求参数

在发送 HTTP 请求前，先验证参数：

```python
# 在代码节点中
excel_path = inputs.get('excel_path', '')
sheet_name = inputs.get('sheet_name', '')
serial = inputs.get('serial', '')

# 验证
if not excel_path:
    output = {"success": False, "error": "excel_path 为空"}
elif not sheet_name:
    output = {"success": False, "error": "sheet_name 为空"}
elif not serial:
    output = {"success": False, "error": "serial 为空"}
else:
    # 发送 HTTP 请求
    output = {"success": True, "ready": True}
```

---

### 方案 3：使用测试接口

先测试服务是否正常工作：

```bash
python test_dify_local_service.py
```

如果测试失败，查看具体错误信息。

---

## 📝 正确的请求格式

### HTTP 请求节点配置

**请求体**：
```json
{
  "excel_path": "C:\\Users\\FH\\source\\repos\\Auto Finan\\Auto Finan\\420财务050823.xlsx",
  "sheet_name": "3-报销",
  "serial": "1"
}
```

**注意**：
- Windows 路径使用 `\\` 或 `/`
- 序号使用字符串 `"1"`，不是数字 `1`
- 工作表名称必须完全匹配

---

## 🐛 常见错误和解决

### 错误 1：文件不存在

**错误信息**：
```json
{
  "error": "Excel 文件不存在: ..."
}
```

**解决**：
1. 检查文件路径是否正确
2. 确保文件在本地服务可以访问的位置
3. 如果 Dify 在远程，文件必须在本地服务器上

---

### 错误 2：序号不存在

**错误信息**：
```json
{
  "error": "未能生成序号 X 的 MCP 提示词"
}
```

**解决**：
1. 检查 Excel 中是否存在该序号
2. 查看 Excel 文件，确认序号格式
3. 尝试其他序号

---

### 错误 3：工作表不存在

**错误信息**：
```json
{
  "error": "处理失败: ..."
}
```

**解决**：
1. 检查工作表名称是否正确
2. 查看 Excel 文件，确认工作表名称
3. 注意大小写和空格

---

## 🔧 调试工具

### 1. 添加详细日志

在服务代码中添加日志输出，查看具体错误。

### 2. 使用测试脚本

```bash
python test_dify_local_service.py
```

### 3. 在 Dify 中查看完整响应

查看 HTTP 请求节点的响应，包含详细错误信息。

---

## ✅ 验证清单

- [ ] Excel 文件路径正确
- [ ] 文件存在于指定路径
- [ ] 工作表名称正确
- [ ] 序号存在于 Excel 中
- [ ] 请求参数格式正确
- [ ] 服务正在运行

---

## 💡 推荐做法

1. **先测试本地**：使用测试脚本验证服务
2. **添加调试**：在代码节点中输出变量值
3. **查看日志**：检查服务日志了解详细错误
4. **验证数据**：确认 Excel 文件和数据正确

---

## 📚 相关文件

- `dify_local_service_flexible.py` - 灵活版本服务（已改进错误处理）
- `test_dify_local_service.py` - 测试脚本
- `Dify_422错误解决方案.md` - 422 错误解决

