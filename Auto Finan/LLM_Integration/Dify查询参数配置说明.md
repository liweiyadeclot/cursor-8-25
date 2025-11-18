# Dify 查询参数配置说明

## ✅ 问题已解决

Dify 将参数放在 URL 查询字符串中，服务已更新支持这种方式。

---

## 🔄 重启服务

**重要**：必须重启服务以应用更改

```bash
# 停止当前服务（Ctrl+C）
# 然后重新启动
cd "Auto Finan\LLM_Integration"
python dify_local_service_flexible.py
```

或使用启动脚本：
```bash
start_dify_local_service.bat
```

---

## 📝 Dify HTTP 请求节点配置

### 当前配置（查询参数方式）

Dify 当前使用查询参数，这是**正确的**，服务已支持。

**配置**：

| 项目 | 值 |
|------|-----|
| 方法 | POST |
| URL | `http://192.168.137.133:8001/api/excel-to-prompt?excel_path={{#workflow.excel_path#}}&sheet_name={{#workflow.sheet_name#}}&serial={{#workflow.serial#}}` |
| 请求头 | `{"Content-Type": "application/json"}`（可选） |
| 请求体 | （留空） |

**注意**：
- 参数在 URL 中，用 `&` 连接
- Dify 会自动进行 URL 编码
- 服务会自动解码

---

### 备选配置（请求体方式）

如果你想使用请求体方式（更标准）：

**配置**：

| 项目 | 值 |
|------|-----|
| 方法 | POST |
| URL | `http://192.168.137.133:8001/api/excel-to-prompt` |
| 请求头 | `{"Content-Type": "application/json"}` |
| 请求体 | `{"excel_path": "{{#workflow.excel_path#}}", "sheet_name": "{{#workflow.sheet_name#}}", "serial": "{{#workflow.serial#}}"}` |

---

## ✅ 验证

重启服务后，测试：

```bash
python test_dify_local_service.py
```

或直接在 Dify 中运行工作流。

---

## 💡 说明

服务现在支持两种方式：
1. **查询参数**（Dify 当前使用）✅
2. **请求体**（标准方式）✅

两种方式都可以工作，无需修改 Dify 配置。

---

## 📚 相关文件

- `dify_local_service_flexible.py` - 已更新，支持查询参数
- `Dify查询参数问题解决.md` - 详细说明

