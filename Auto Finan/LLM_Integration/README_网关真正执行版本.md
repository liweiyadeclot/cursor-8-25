# Playwright MCP HTTP 网关（真正执行版本）

## 🎉 新功能

网关现在可以**真正执行浏览器操作**，而不仅仅是解析提示词！

---

## 🚀 快速开始

### 启动服务

```bash
cd "Auto Finan\LLM_Integration"
start_mcp_gateway.bat
```

或直接运行：

```bash
python playwright_mcp_http_gateway_executor.py
```

---

## ✨ 功能特点

### 支持的操作

1. ✅ **打开页面** - `打开https://example.com`
2. ✅ **输入文本** - `在用户名输入框中输入test`
3. ✅ **下拉框选择** - `在支付方式下拉框中选择值为"个人转卡"`
4. ✅ **点击按钮** - `点击登录按钮`
5. ✅ **填写表单** - `向电费输入框填写100`
6. ✅ **选择日期** - `选择日期起始时间为2024-12-26`
7. ✅ **等待操作** - `等待页面响应`
8. ✅ **保存图片** - `将验证码图片保存至...`
9. ✅ **运行脚本** - `运行OCR.py`
10. ⚠️ **调用脚本（带参数）** - 部分支持
11. ⚠️ **文件重命名** - 部分支持

---

## 📝 使用示例

### HTTP 请求

```bash
curl -X POST http://localhost:3030/mcp/execute \
  -H "Content-Type: application/json" \
  -d '{
    "prompt": "1. 打开https://example.com\n2. 在搜索输入框中输入test\n3. 点击搜索按钮",
    "headless": false
  }'
```

### Dify 配置

**HTTP 请求节点**：

| 项目 | 值 |
|------|-----|
| **方法** | `POST` |
| **URL** | `http://localhost:3030/mcp/execute` |
| **请求头** | `{"Content-Type": "application/json"}` |
| **请求体** | `{"prompt": "{{#workflow.mcp_prompt#}}", "headless": false}` |

---

## ⚙️ 配置选项

### 请求参数

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `prompt` | string | 必填 | MCP 提示词 |
| `timeout` | int | 300 | 总超时时间（秒） |
| `browser` | string | "chrome" | 浏览器类型（chrome/firefox/webkit） |
| `headless` | bool | false | 是否无头模式（false=显示浏览器） |

---

## 🔧 技术实现

### 执行流程

1. **解析提示词** - 提取操作步骤
2. **启动浏览器** - 使用 Playwright
3. **执行步骤** - 逐个执行操作
4. **记录日志** - 记录每个步骤的执行情况
5. **返回结果** - 返回执行结果和日志

### 选择器策略

使用多种选择器策略，提高成功率：

1. Label 文本匹配
2. Placeholder 匹配
3. Name 属性匹配
4. ID 属性匹配
5. XPath 选择器

---

## 📊 响应格式

### 成功响应

```json
{
  "status": "success",
  "message": "所有步骤执行成功（53/53）",
  "execution_id": "exec_20251118_144930",
  "logs": [
    "🚀 开始执行，共 53 个步骤",
    "✅ 浏览器已启动: chromium (headless=false)",
    "--- 步骤 1/53: 打开https://... ---",
    "✅ 页面已打开: https://...",
    ...
  ],
  "total_steps": 53,
  "success_steps": 53,
  "failed_steps": [],
  "timestamp": "2025-11-18T14:49:30.356111"
}
```

### 部分成功响应

```json
{
  "status": "partial",
  "message": "部分步骤执行成功（50/53）",
  "execution_id": "exec_20251118_144930",
  "logs": [...],
  "total_steps": 53,
  "success_steps": 50,
  "failed_steps": [
    [3, "在验证码输入框中输入1234"]
  ],
  "timestamp": "2025-11-18T14:49:30.356111"
}
```

---

## ⚠️ 注意事项

### 1. 浏览器显示

- 默认 `headless=false`，会显示浏览器窗口
- 可以设置 `headless=true` 使用无头模式

### 2. 超时设置

- 默认总超时时间：300 秒（5分钟）
- 每个步骤的超时时间：5 秒
- 可以根据需要调整

### 3. 特殊操作

某些操作需要特殊处理：
- **调用脚本（带参数）** - 需要解析参数并传递
- **文件重命名** - 需要文件系统操作
- **OCR 识别** - 需要集成 OCR 功能

---

## 🔍 调试

### 查看日志

执行结果中的 `logs` 字段包含详细的执行日志：

```python
result = execute_playwright_prompt(prompt)
for log in result["logs"]:
    print(log)
```

### 失败步骤

如果某些步骤失败，检查 `failed_steps` 字段：

```python
if result["failed_steps"]:
    for step_num, step in result["failed_steps"]:
        print(f"步骤 {step_num} 失败: {step}")
```

---

## 📚 相关文件

- `playwright_mcp_http_gateway_executor.py` - 真正执行版本
- `playwright_mcp_http_gateway.py` - 原版本（仅解析）
- `playwright_direct_executor.py` - 直接执行器（独立版本）

---

## 🎯 对比

| 版本 | 功能 | 适用场景 |
|------|------|---------|
| **原版本** | 仅解析提示词 | 测试、验证格式 |
| **执行版本** | 真正执行操作 | 生产环境、自动化 |

---

## ✅ 总结

现在网关可以：
- ✅ 真正打开浏览器
- ✅ 执行所有操作步骤
- ✅ 返回详细的执行日志
- ✅ 支持多种浏览器
- ✅ 支持有头/无头模式

**重启服务后，Dify 工作流就可以真正执行浏览器操作了！**

