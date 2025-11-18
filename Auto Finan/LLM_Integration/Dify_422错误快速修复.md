# Dify 422 错误快速修复

## ❌ 错误：Request failed with status code 422

**422 错误** = 请求体验证失败

---

## ✅ 快速修复（3 步）

### 步骤 1：重启服务（使用灵活版本）

```bash
# 停止当前服务（Ctrl+C）
# 然后启动灵活版本
cd "Auto Finan\LLM_Integration"
python dify_local_service_flexible.py
```

或使用启动脚本（已自动使用灵活版本）：
```bash
start_dify_local_service.bat
```

---

### 步骤 2：检查 Dify HTTP 请求节点配置

**请求体必须是**：
```json
{
  "excel_path": "{{#workflow.excel_path#}}",
  "sheet_name": "{{#workflow.sheet_name#}}",
  "serial": "{{#workflow.serial#}}"
}
```

**检查点**：
- ✅ 使用双引号 `"`，不是单引号
- ✅ 字段名：`excel_path`（下划线），不是 `excelPath`
- ✅ 所有三个字段都存在
- ✅ Content-Type: `application/json`

---

### 步骤 3：验证变量值

在代码节点中添加调试输出：

```python
# 在生成提示词的代码节点中
output = {
    "debug": {
        "excel_path": inputs.get('excel_path', ''),
        "sheet_name": inputs.get('sheet_name', ''),
        "serial": inputs.get('serial', '')
    }
}
```

查看输出，确认变量值是否正确。

---

## 🔍 常见原因

### 原因 1：字段名错误

❌ **错误**：
```json
{
  "excelPath": "...",
  "sheetName": "..."
}
```

✅ **正确**：
```json
{
  "excel_path": "...",
  "sheet_name": "..."
}
```

---

### 原因 2：变量未定义

如果 Dify 变量不存在，会传递字符串 `"{{#workflow.excel_path#}}"`，导致验证失败。

**解决**：检查上游节点是否正确输出变量。

---

### 原因 3：请求体格式错误

❌ **错误**：不是 JSON 格式
```
excel_path={{#workflow.excel_path#}}
```

✅ **正确**：JSON 格式
```json
{
  "excel_path": "{{#workflow.excel_path#}}"
}
```

---

## 📝 正确的 Dify 配置

### HTTP 请求节点

| 配置项 | 值 |
|--------|-----|
| **方法** | POST |
| **URL** | `http://192.168.137.133:8001/api/excel-to-prompt` |
| **请求头** | `{"Content-Type": "application/json"}` |
| **请求体** | 见下方 |
| **超时** | 60 秒 |

**请求体（JSON）**：
```json
{
  "excel_path": "{{#workflow.excel_path#}}",
  "sheet_name": "{{#workflow.sheet_name#}}",
  "serial": "{{#workflow.serial#}}"
}
```

---

## 🚀 使用灵活版本服务

灵活版本服务（`dify_local_service_flexible.py`）提供：
- ✅ 更详细的错误信息
- ✅ 更好的字段验证
- ✅ 支持多种请求格式

**启动**：
```bash
python dify_local_service_flexible.py
```

---

## ✅ 验证

运行测试：
```bash
python test_dify_local_service.py
```

如果测试通过，说明服务正常，问题在 Dify 配置。

---

## 📚 详细文档

- `Dify_422错误解决方案.md` - 详细解决方案
- `dify_local_service_flexible.py` - 灵活版本服务

