# Playwright MCP HTTP 调用示例

## 📋 概述

本文档展示如何通过 HTTP 方式调用 Playwright MCP，适用于 Dify 工作流或其他 HTTP 客户端。

---

## 📤 请求格式

### HTTP 请求示例

**请求方法：** `POST`  
**请求头：** `Content-Type: application/json`  
**请求体：** JSON 格式

```json
{
  "prompt": "1. 请你调用Playwright MCP，执行以下命令，一次性执行完\n2. 打开https://cwcx.uestc.edu.cn/WFManager/login.jsp\n3. 业务大类：报销业务。以下是需要执行的页面操作：\n4. 在用户名输入框中输入5130008\n5. 在密码输入框中输入Uestc418\n6. 将验证码图片保存至C:\\Users\\FH\\source\\repos\\Auto Finan\\Auto Finan\\LLM_Integration目录下，命名为example.jpg\n7. 读取图片中的验证码信息\n8. 输入验证码\n9. 点击登录按钮\n10. 点击网上预约报账按钮\n11. 点击申请报销单按钮\n12. 点击已阅读并同意按钮\n13. 在报销项目号输入框中输入M112023ZHCG0006\n14. 在附件张数输入框中输入2\n15. 在支付方式下拉框中选择个人转卡\n16. 点击下一步按钮\n17. 向专利费输入框填写100\n18. 点击下一步按钮\n19. 在学工号输入框中输入5070016\n20. 银行卡号尾号内容为2818\n21. 在金额输入框中输入50\n22. 点击提交按钮\n23. 等待页面响应\n24. 在学工号输入框中输入202422090507\n25. 银行卡号尾号内容为5054\n26. 在金额输入框中输入50\n27. 点击下一步按钮\n28. 选择日期预约日期为2025-10-20\n29. 点击预约按钮\n30. 点击打印确认单按钮\n31. 调用test_mouse_keyboard.py，执行一个python自动点击的脚本，脚本的第一个参数为保存路径，第二个参数为保存文件名，请你以当前页面中的信息，以报销单号-项目号-金额的格式，输入第二个参数\n32. 等待刚刚运行的脚本运行完毕\n33. 点击返回按钮\n34. 重命名当前读取的提示词文件，将未预约改成已预约"
}
```

### 简化版请求（仅核心操作）

```json
{
  "prompt": "请你调用Playwright MCP，执行以下命令，一次性执行完\n打开https://cwcx.uestc.edu.cn/WFManager/login.jsp\n业务大类：报销业务。以下是需要执行的页面操作：\n在用户名输入框中输入5130008\n在密码输入框中输入Uestc418\n将验证码图片保存至C:\\Users\\FH\\source\\repos\\Auto Finan\\Auto Finan\\LLM_Integration目录下，命名为example.jpg\n读取图片中的验证码信息\n输入验证码\n点击登录按钮\n点击网上预约报账按钮\n点击申请报销单按钮\n点击已阅读并同意按钮\n在报销项目号输入框中输入M112023ZHCG0006\n在附件张数输入框中输入2\n在支付方式下拉框中选择个人转卡\n点击下一步按钮\n向专利费输入框填写100\n点击下一步按钮\n在学工号输入框中输入5070016\n银行卡号尾号内容为2818\n在金额输入框中输入50\n点击提交按钮\n等待页面响应\n在学工号输入框中输入202422090507\n银行卡号尾号内容为5054\n在金额输入框中输入50\n点击下一步按钮\n选择日期预约日期为2025-10-20\n点击预约按钮\n点击打印确认单按钮"
}
```

---

## 📥 响应格式

### 成功响应示例

```json
{
  "status": "success",
  "message": "Playwright MCP 执行完成",
  "execution_id": "exec_20250110_143025",
  "logs": [
    "已打开页面: https://cwcx.uestc.edu.cn/WFManager/login.jsp",
    "已输入用户名: 5130008",
    "已输入密码: ***",
    "验证码已识别并输入",
    "登录成功",
    "已点击网上预约报账按钮",
    "..."
  ],
  "screenshots": [
    "screenshot_1.png",
    "screenshot_2.png"
  ],
  "timestamp": "2025-01-10T14:30:25Z"
}
```

### 错误响应示例

```json
{
  "status": "error",
  "error_code": "MCP_EXECUTION_FAILED",
  "message": "执行失败：无法找到元素 '登录按钮'",
  "error_details": {
    "step": 9,
    "action": "点击登录按钮",
    "error": "Element not found: button[type='submit']"
  },
  "timestamp": "2025-01-10T14:30:25Z"
}
```

### 部分成功响应示例

```json
{
  "status": "partial",
  "message": "部分步骤执行成功",
  "completed_steps": 15,
  "total_steps": 34,
  "failed_at_step": 16,
  "error": "无法找到元素 '下一步按钮'",
  "logs": [
    "已打开页面: https://cwcx.uestc.edu.cn/WFManager/login.jsp",
    "...",
    "已点击已阅读并同意按钮"
  ],
  "timestamp": "2025-01-10T14:30:25Z"
}
```

---

## 🔧 使用 cURL 测试

### 基本调用

```bash
curl -X POST http://localhost:3030/mcp/execute \
  -H "Content-Type: application/json" \
  -d "{\"prompt\": \"请你调用Playwright MCP，执行以下命令，一次性执行完\\n打开https://cwcx.uestc.edu.cn/WFManager/login.jsp\\n业务大类：报销业务。以下是需要执行的页面操作：\\n在用户名输入框中输入5130008\\n在密码输入框中输入Uestc418\"}"
```

### 从文件读取提示词

```bash
# 读取提示词文件
PROMPT=$(cat "Auto Finan/LLM_Integration/mcp_prompts/未预约-M112023ZHCG0006-100-20251010-08-20-08.txt")

# 发送请求（Windows PowerShell）
$prompt = Get-Content "Auto Finan\LLM_Integration\mcp_prompts\未预约-M112023ZHCG0006-100-20251010-08-20-08.txt" -Raw
$body = @{ prompt = $prompt } | ConvertTo-Json
Invoke-RestMethod -Uri "http://localhost:3030/mcp/execute" -Method Post -Body $body -ContentType "application/json"
```

---

## 🎯 在 Dify 工作流中使用

### 方法 1：使用 HTTP 请求节点

1. **添加 HTTP 请求节点**
   - 节点类型：`HTTP 请求`
   - 请求方法：`POST`
   - URL：`http://localhost:3030/mcp/execute`（或你的 MCP 网关地址）

2. **配置请求体**
   ```json
   {
     "prompt": "{{#workflow.prompt#}}"
   }
   ```
   其中 `{{#workflow.prompt#}}` 是上游节点生成的提示词变量。

3. **从上游节点获取提示词**
   - 如果上游是代码节点，可以直接生成提示词字符串
   - 如果上游是 LLM 节点，需要提取生成的文本内容

### 方法 2：使用代码节点生成并调用

在 Dify 的代码节点中：

```python
import requests
import json

# 从上游节点获取数据（例如从 Excel 处理节点）
excel_data = {{#workflow.excel_data#}}

# 生成 MCP 提示词（这里简化示例，实际应调用你的 workflow_core）
mcp_prompt = f"""请你调用Playwright MCP，执行以下命令，一次性执行完
打开https://cwcx.uestc.edu.cn/WFManager/login.jsp
业务大类：报销业务。以下是需要执行的页面操作：
在用户名输入框中输入{excel_data.get('username', '')}
在密码输入框中输入{excel_data.get('password', '')}
..."""

# 调用 MCP HTTP 接口
mcp_endpoint = "http://localhost:3030/mcp/execute"
response = requests.post(
    mcp_endpoint,
    json={"prompt": mcp_prompt},
    timeout=300  # 5分钟超时
)

result = response.json()
print(f"执行状态: {result.get('status')}")
print(f"执行日志: {result.get('logs', [])}")

# 返回结果给下游节点
return {
    "status": result.get("status"),
    "message": result.get("message"),
    "logs": result.get("logs", [])
}
```

### 方法 3：从文件读取提示词

如果提示词已保存在文件中：

```python
import requests
import os

# 读取提示词文件
prompt_file = "Auto Finan/LLM_Integration/mcp_prompts/未预约-M112023ZHCG0006-100-20251010-08-20-08.txt"
with open(prompt_file, 'r', encoding='utf-8') as f:
    prompt = f.read()

# 调用 MCP
response = requests.post(
    "http://localhost:3030/mcp/execute",
    json={"prompt": prompt},
    timeout=300
)

return response.json()
```

---

## 📝 Python 完整示例

### 示例 1：直接调用

```python
import requests
import json

def call_playwright_mcp_via_http(prompt: str, endpoint: str = "http://localhost:3030/mcp/execute") -> dict:
    """通过 HTTP 调用 Playwright MCP"""
    try:
        response = requests.post(
            endpoint,
            json={"prompt": prompt},
            timeout=300,  # 5分钟超时
            headers={"Content-Type": "application/json"}
        )
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        return {
            "status": "error",
            "message": f"HTTP 请求失败: {str(e)}"
        }

# 使用示例
prompt = """请你调用Playwright MCP，执行以下命令，一次性执行完
打开https://cwcx.uestc.edu.cn/WFManager/login.jsp
业务大类：报销业务。以下是需要执行的页面操作：
在用户名输入框中输入5130008
在密码输入框中输入Uestc418"""

result = call_playwright_mcp_via_http(prompt)
print(json.dumps(result, indent=2, ensure_ascii=False))
```

### 示例 2：从 Excel 生成并调用

```python
import requests
from workflow_core import process_excel_to_mcp_direct

# 从 Excel 生成 MCP 提示词
excel_path = r"C:\Users\FH\source\repos\Auto Finan\Auto Finan\420财务050823.xlsx"
sheet_name = "3-报销"
serial = "1"

mcp_prompt = process_excel_to_mcp_direct(excel_path, sheet_name, serial)

if mcp_prompt:
    # 调用 MCP HTTP 接口
    response = requests.post(
        "http://localhost:3030/mcp/execute",
        json={"prompt": mcp_prompt},
        timeout=300
    )
    
    result = response.json()
    print(f"执行状态: {result.get('status')}")
    if result.get('status') == 'success':
        print("✅ 自动化执行成功")
    else:
        print(f"❌ 执行失败: {result.get('message')}")
```

---

## 🔍 提示词格式说明

### 标准格式

提示词是一个多行文本字符串，每行代表一个操作步骤：

```
1. 请你调用Playwright MCP，执行以下命令，一次性执行完
2. 打开<URL>
3. 业务大类：<业务类型>。以下是需要执行的页面操作：
4. <操作步骤1>
5. <操作步骤2>
...
```

### 操作步骤类型

- **输入操作**：`在{控件名}输入框中输入{值}`
- **下拉选择**：`在{控件名}下拉框中选择{值}`
- **点击操作**：`点击{按钮名}按钮`
- **日期选择**：`选择日期{控件名}为{日期}`
- **文件操作**：`将验证码图片保存至{路径}，命名为{文件名}`
- **脚本调用**：`调用{脚本名}，执行{说明}`

---

## ⚙️ 环境变量配置

在调用前设置环境变量：

```bash
# Windows CMD
set MCP_HTTP_ENDPOINT=http://localhost:3030/mcp/execute

# Windows PowerShell
$env:MCP_HTTP_ENDPOINT="http://localhost:3030/mcp/execute"

# Linux/Mac
export MCP_HTTP_ENDPOINT=http://localhost:3030/mcp/execute
```

---

## 🚨 注意事项

1. **超时设置**：自动化操作可能需要较长时间，建议设置 5-10 分钟超时
2. **错误处理**：检查响应中的 `status` 字段，处理 `error` 和 `partial` 状态
3. **日志记录**：保存响应中的 `logs` 数组，便于调试和审计
4. **并发控制**：避免同时执行多个自动化任务，可能导致浏览器冲突
5. **路径格式**：Windows 路径中的反斜杠需要转义或使用双反斜杠

---

## 📚 相关文件

- `run_excel_reimburse_workflow.py` - 包含 HTTP 调用示例代码
- `workflow_core.py` - MCP 提示词生成逻辑
- `excel_batch_processor.py` - 批量生成提示词文件

