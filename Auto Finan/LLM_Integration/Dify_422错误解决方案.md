# Dify 422 错误解决方案

## ❌ 错误：Request failed with status code 422

**422 错误**表示请求体验证失败，通常是字段格式或类型问题。

---

## 🔍 问题原因

在 FastAPI 中，422 错误通常是因为：

1. **字段缺失**：缺少必需字段
2. **字段类型错误**：字段类型不匹配
3. **字段名错误**：字段名拼写错误
4. **请求体格式错误**：JSON 格式不正确

---

## ✅ 解决方案

### 方案 1：检查 Dify HTTP 请求节点配置

#### 请求体格式

**正确格式**：
```json
{
  "excel_path": "C:\\Users\\FH\\source\\repos\\Auto Finan\\Auto Finan\\420财务050823.xlsx",
  "sheet_name": "3-报销",
  "serial": "1"
}
```

**常见错误**：

❌ **错误 1**：字段名错误
```json
{
  "excelPath": "...",  // ❌ 应该是 excel_path
  "sheetName": "...",  // ❌ 应该是 sheet_name
  "serial": "1"
}
```

❌ **错误 2**：缺少字段
```json
{
  "excel_path": "..."
  // ❌ 缺少 sheet_name 和 serial
}
```

❌ **错误 3**：使用模板语法错误
```json
{
  "excel_path": "{{#workflow.excel_path#}}",  // ✅ 正确
  "sheet_name": "{{#workflow.sheet_name#}}",  // ✅ 正确
  "serial": "{{#workflow.serial#}}"           // ✅ 正确
}
```

---

### 方案 2：使用灵活版本服务（推荐）

我已经创建了灵活版本的服务，支持更宽松的验证：

**启动灵活版本**：

```bash
cd "Auto Finan\LLM_Integration"
python dify_local_service_flexible.py
```

**优点**：
- ✅ 更详细的错误信息
- ✅ 支持多种字段名格式
- ✅ 更好的错误处理

---

### 方案 3：检查 Dify 变量传递

#### 在 Dify HTTP 请求节点中

**请求体配置**：

```json
{
  "excel_path": "{{#workflow.excel_path#}}",
  "sheet_name": "{{#workflow.sheet_name#}}",
  "serial": "{{#workflow.serial#}}"
}
```

**检查变量值**：

1. 在代码节点中输出变量：
   ```python
   output = {
       "debug_excel_path": inputs.get('excel_path', ''),
       "debug_sheet_name": inputs.get('sheet_name', ''),
       "debug_serial": inputs.get('serial', '')
   }
   ```

2. 查看输出，确认变量值是否正确

---

## 🔧 调试步骤

### 步骤 1：测试服务接口

使用测试脚本：

```bash
python test_dify_local_service.py
```

### 步骤 2：查看服务日志

服务启动后会显示请求日志，查看具体错误信息。

### 步骤 3：使用 Postman 测试

**请求配置**：
- URL: `http://localhost:8001/api/excel-to-prompt`
- Method: `POST`
- Headers: `Content-Type: application/json`
- Body:
```json
{
  "excel_path": "C:\\Users\\FH\\source\\repos\\Auto Finan\\Auto Finan\\420财务050823.xlsx",
  "sheet_name": "3-报销",
  "serial": "1"
}
```

---

## 📝 正确的 Dify 配置

### HTTP 请求节点配置

**节点类型**：`HTTP 请求`

**配置**：

| 项目 | 值 |
|------|-----|
| 方法 | POST |
| URL | `http://192.168.137.133:8001/api/excel-to-prompt` |
| 请求头 | `{"Content-Type": "application/json"}` |
| 请求体 | 见下方 |
| 超时 | 60 秒 |

**请求体（JSON）**：
```json
{
  "excel_path": "{{#workflow.excel_path#}}",
  "sheet_name": "{{#workflow.sheet_name#}}",
  "serial": "{{#workflow.serial#}}"
}
```

**重要**：
- 使用双引号 `"`，不是单引号 `'`
- 字段名使用下划线：`excel_path`，不是驼峰：`excelPath`
- 确保所有三个字段都存在

---

## 🐛 常见错误和解决

### 错误 1：字段名错误

**错误**：
```json
{
  "excelPath": "...",  // ❌
  "sheetName": "..."   // ❌
}
```

**解决**：使用下划线格式
```json
{
  "excel_path": "...",  // ✅
  "sheet_name": "..."   // ✅
}
```

### 错误 2：变量未定义

**错误**：
```json
{
  "excel_path": "{{#workflow.excel_path#}}",  // 如果变量不存在，会传递字符串 "{{#workflow.excel_path#}}"
}
```

**解决**：
1. 检查上游节点是否正确输出变量
2. 检查变量名是否匹配
3. 在代码节点中添加调试输出

### 错误 3：字段类型错误

**错误**：
```json
{
  "serial": 1  // ❌ 数字类型
}
```

**解决**：使用字符串
```json
{
  "serial": "1"  // ✅ 字符串类型
}
```

---

## ✅ 验证清单

- [ ] 请求体是有效的 JSON
- [ ] 字段名正确：`excel_path`, `sheet_name`, `serial`
- [ ] 所有三个字段都存在
- [ ] 字段值不是空字符串
- [ ] Content-Type 头设置为 `application/json`
- [ ] 服务正在运行

---

## 💡 推荐做法

1. **使用灵活版本服务**：`dify_local_service_flexible.py`
2. **添加调试输出**：在代码节点中输出变量值
3. **测试接口**：使用 Postman 或测试脚本验证
4. **查看日志**：检查服务日志了解详细错误

---

## 📚 相关文件

- `dify_local_service.py` - 标准版本
- `dify_local_service_flexible.py` - 灵活版本（推荐）
- `test_dify_local_service.py` - 测试脚本
- `Dify连接问题解决方案.md` - 连接问题解决

