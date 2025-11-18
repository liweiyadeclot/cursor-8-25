# Dify HTTP 请求体配置

## 📝 JSON 请求体配置

### 方式 1：标准 JSON 格式（推荐）

在 Dify HTTP 请求节点的**请求体**中，使用以下 JSON：

```json
{
  "excel_path": "{{#workflow.excel_path#}}",
  "sheet_name": "{{#workflow.sheet_name#}}",
  "serial": "{{#workflow.serial#}}"
}
```

---

### 方式 2：带默认值的格式

如果担心变量为空，可以添加默认值：

```json
{
  "excel_path": "{{#workflow.excel_path#}}",
  "sheet_name": "{{#workflow.sheet_name#}}",
  "serial": "{{#workflow.serial#}}",
  "timeout": 300
}
```

---

## 🔧 完整 HTTP 请求节点配置

### 节点类型：HTTP 请求

**配置项**：

| 配置项 | 值 |
|--------|-----|
| **方法** | `POST` |
| **URL** | `http://192.168.137.133:8001/api/excel-to-prompt`<br>（替换为你的实际 IP） |
| **请求头** | 见下方 |
| **请求体** | 见下方 |
| **超时** | `60` 秒（或更长） |

---

### 请求头配置

**格式**：JSON

```json
{
  "Content-Type": "application/json"
}
```

---

### 请求体配置

**格式**：JSON

```json
{
  "excel_path": "{{#workflow.excel_path#}}",
  "sheet_name": "{{#workflow.sheet_name#}}",
  "serial": "{{#workflow.serial#}}"
}
```

---

## 📋 完整示例

### 示例 1：基本配置

**HTTP 请求节点**：

- **方法**：`POST`
- **URL**：`http://192.168.137.133:8001/api/excel-to-prompt`
- **请求头**：
```json
{
  "Content-Type": "application/json"
}
```
- **请求体**：
```json
{
  "excel_path": "{{#workflow.excel_path#}}",
  "sheet_name": "{{#workflow.sheet_name#}}",
  "serial": "{{#workflow.serial#}}"
}
```

---

### 示例 2：带调试信息

如果需要调试，可以在请求体中添加额外字段：

```json
{
  "excel_path": "{{#workflow.excel_path#}}",
  "sheet_name": "{{#workflow.sheet_name#}}",
  "serial": "{{#workflow.serial#}}",
  "debug": true
}
```

（服务会忽略 `debug` 字段，但可以用于调试）

---

## ⚠️ 注意事项

### 1. 变量引用格式

**正确**：
```json
{
  "excel_path": "{{#workflow.excel_path#}}"
}
```

**错误**：
```json
{
  "excel_path": {{#workflow.excel_path#}}  // ❌ 缺少引号
}
```

```json
{
  "excel_path": "#workflow.excel_path#"  // ❌ 缺少 {{ }}
}
```

---

### 2. JSON 格式

- 使用双引号 `"`，不是单引号 `'`
- 最后一个字段后不要加逗号
- 确保 JSON 格式有效

---

### 3. 字段名

- 使用下划线：`excel_path`，不是驼峰：`excelPath`
- 字段名必须完全匹配

---

## 🔍 验证配置

### 在代码节点中验证变量

在发送 HTTP 请求前，添加验证节点：

```python
excel_path = inputs.get('excel_path', '')
sheet_name = inputs.get('sheet_name', '')
serial = inputs.get('serial', '')

# 验证并输出
output = {
    "excel_path": excel_path,
    "sheet_name": sheet_name,
    "serial": serial,
    "all_valid": bool(excel_path and sheet_name and serial)
}
```

---

## 📝 复制粘贴模板

直接复制以下内容到 Dify HTTP 请求节点的请求体：

```json
{
  "excel_path": "{{#workflow.excel_path#}}",
  "sheet_name": "{{#workflow.sheet_name#}}",
  "serial": "{{#workflow.serial#}}"
}
```

---

## 🆚 两种方式对比

### 方式 1：查询参数（Dify 当前使用）

**URL**：
```
http://192.168.137.133:8001/api/excel-to-prompt?excel_path={{#workflow.excel_path#}}&sheet_name={{#workflow.sheet_name#}}&serial={{#workflow.serial#}}
```

**请求体**：留空

**优点**：
- Dify 默认使用这种方式
- 简单直接

**缺点**：
- URL 可能很长
- 参数暴露在 URL 中

---

### 方式 2：请求体（推荐）

**URL**：
```
http://192.168.137.133:8001/api/excel-to-prompt
```

**请求体**：
```json
{
  "excel_path": "{{#workflow.excel_path#}}",
  "sheet_name": "{{#workflow.sheet_name#}}",
  "serial": "{{#workflow.serial#}}"
}
```

**优点**：
- 更标准
- 更安全
- 支持更复杂的数据

**缺点**：
- 需要配置请求体

---

## ✅ 推荐配置

**推荐使用请求体方式**（方式 2），因为：
- ✅ 更标准
- ✅ 更安全
- ✅ 更灵活

**配置**：

| 项目 | 值 |
|------|-----|
| 方法 | POST |
| URL | `http://192.168.137.133:8001/api/excel-to-prompt` |
| 请求头 | `{"Content-Type": "application/json"}` |
| 请求体 | `{"excel_path": "{{#workflow.excel_path#}}", "sheet_name": "{{#workflow.sheet_name#}}", "serial": "{{#workflow.serial#}}"}` |

---

## 📚 相关文件

- `dify_local_service_flexible.py` - 支持两种方式
- `Dify使用本地服务配置.md` - 完整配置指南

