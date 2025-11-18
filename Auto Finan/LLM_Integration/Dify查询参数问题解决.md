# Dify 查询参数问题解决

## ❌ 问题：请求体不是有效的 JSON

**原因**：Dify 将参数放在 URL 查询字符串中，而不是请求体中。

**Dify 发送的请求**：
```
POST /api/excel-to-prompt?excel_path=...&sheet_name=...&serial=1
```

**服务期望的请求**：
```
POST /api/excel-to-prompt
Content-Type: application/json

{
  "excel_path": "...",
  "sheet_name": "...",
  "serial": "1"
}
```

---

## ✅ 解决方案

### 已修复

我已经更新了 `dify_local_service_flexible.py`，现在支持：

1. **从查询参数读取**（Dify 的方式）
2. **从请求体读取**（标准方式）
3. **自动 URL 解码**（处理 URL 编码的参数）

---

## 🔄 重启服务

**重要**：需要重启服务以应用更改

```bash
# 停止当前服务（Ctrl+C）
# 然后重新启动
python dify_local_service_flexible.py
```

或使用启动脚本：
```bash
start_dify_local_service.bat
```

---

## 📝 Dify 配置（两种方式都可以）

### 方式 1：使用查询参数（Dify 当前方式）

**HTTP 请求节点配置**：

| 项目 | 值 |
|------|-----|
| 方法 | POST |
| URL | `http://192.168.137.133:8001/api/excel-to-prompt?excel_path={{#workflow.excel_path#}}&sheet_name={{#workflow.sheet_name#}}&serial={{#workflow.serial#}}` |
| 请求头 | `{"Content-Type": "application/json"}` |
| 请求体 | （留空或删除） |

**注意**：URL 中的参数需要 URL 编码，Dify 会自动处理。

---

### 方式 2：使用请求体（推荐）

**HTTP 请求节点配置**：

| 项目 | 值 |
|------|-----|
| 方法 | POST |
| URL | `http://192.168.137.133:8001/api/excel-to-prompt` |
| 请求头 | `{"Content-Type": "application/json"}` |
| 请求体 | `{"excel_path": "{{#workflow.excel_path#}}", "sheet_name": "{{#workflow.sheet_name#}}", "serial": "{{#workflow.serial#}}"}` |

---

## 🔍 验证

重启服务后，两种方式都应该可以工作：

1. **查询参数方式**（Dify 当前使用）
2. **请求体方式**（标准方式）

---

## 💡 推荐

**推荐使用方式 2（请求体）**，因为：
- ✅ 更标准
- ✅ 支持更复杂的数据
- ✅ 更安全（参数不在 URL 中）

但如果 Dify 只支持查询参数，方式 1 也可以工作。

---

## 📚 相关文件

- `dify_local_service_flexible.py` - 已更新，支持查询参数
- `start_dify_local_service.bat` - 启动脚本

