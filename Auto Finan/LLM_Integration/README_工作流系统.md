# 工作流系统使用指南

## 📋 概述

工作流系统允许你在 Cursor 中通过配置文件定义自动化任务，无需对话即可自动执行 MCP 操作。

## 🚀 快速开始

### 1. 启动 MCP HTTP 网关

```bash
# 在终端中运行
python playwright_mcp_http_gateway_executor.py
# 或使用批处理文件
start_playwright_mcp_http.bat
```

确保网关运行在 `http://localhost:3030`

### 2. 创建工作流配置文件

创建工作流 JSON 文件，例如 `my_workflow.json`:

```json
{
  "name": "我的工作流",
  "variables": {
    "project_code": "M112023ZHCG0006"
  },
  "steps": [
    {
      "type": "mcp",
      "name": "执行登录和报销",
      "prompt": "1. 打开https://cwcx.uestc.edu.cn/WFManager/login.jsp\n2. 在用户名输入框中输入5130008\n3. 在密码输入框中输入Uestc418\n4. 等待 20 秒\n5. 点击id为zhLogin的登录按钮",
      "timeout": 300,
      "browser": "chrome",
      "headless": false
    }
  ]
}
```

### 3. 执行工作流

在 Cursor 中运行：

```bash
python workflow_engine.py my_workflow.json
```

## 📝 工作流配置格式

### 基本结构

```json
{
  "name": "工作流名称",
  "description": "工作流描述（可选）",
  "variables": {
    "变量名": "变量值"
  },
  "steps": [
    {
      "type": "步骤类型",
      "name": "步骤名称（可选）",
      ...
    }
  ]
}
```

## 🔧 支持的步骤类型

### 1. MCP 调用步骤 (`mcp`)

执行 Playwright MCP 命令：

```json
{
  "type": "mcp",
  "name": "执行登录",
  "prompt": "1. 打开https://example.com\n2. 点击登录按钮",
  "timeout": 300,
  "browser": "chrome",
  "headless": false,
  "session_id": "可选会话ID",
  "save_result_to": "result_var",
  "continue_on_error": false
}
```

**参数说明：**
- `prompt`: MCP 提示词（支持变量 `${var_name}`）
- `timeout`: 超时时间（秒），默认 300
- `browser`: 浏览器类型（chrome/firefox/webkit），默认 chrome
- `headless`: 是否无头模式，默认 false
- `session_id`: 会话ID（用于保持浏览器会话）
- `save_result_to`: 将执行结果保存到变量
- `continue_on_error`: 失败时是否继续，默认 false

### 2. 设置变量步骤 (`set_variable`)

设置工作流变量：

```json
{
  "type": "set_variable",
  "name": "project_code",
  "value": "M112023ZHCG0006"
}
```

### 3. 等待步骤 (`wait`)

暂停执行：

```json
{
  "type": "wait",
  "seconds": 5
}
```

### 4. 条件判断步骤 (`condition`)

根据条件执行不同分支：

```json
{
  "type": "condition",
  "condition": "${login_result.status} == 'success'",
  "if_true": [
    {
      "type": "log",
      "message": "登录成功"
    }
  ],
  "if_false": [
    {
      "type": "log",
      "message": "登录失败"
    }
  ]
}
```

### 5. 循环步骤 (`loop`)

循环执行步骤：

```json
{
  "type": "loop",
  "item_var": "current_item",
  "items": ["1", "2", "3"],
  "steps": [
    {
      "type": "log",
      "message": "处理: ${current_item}"
    }
  ]
}
```

### 6. 日志步骤 (`log`)

输出日志：

```json
{
  "type": "log",
  "message": "当前处理序号: ${serial}"
}
```

### 7. 脚本步骤 (`script`)

执行 Python 脚本：

```json
{
  "type": "script",
  "script": "path/to/script.py",
  "timeout": 60
}
```

### 8. Excel到提示词步骤 (`excel_to_prompt`)

从Excel文件读取数据，生成MCP提示词：

```json
{
  "type": "excel_to_prompt",
  "name": "从Excel生成MCP提示词",
  "excel_path": "C:\\path\\to\\file.xlsx",
  "sheet_name": "3-报销",
  "serial": "1",
  "use_llm": true,
  "save_to": "mcp_prompt"
}
```

**参数说明：**
- `excel_path`: Excel文件路径（支持变量 `${var_name}`）
- `sheet_name`: 工作表名称（可选，支持变量）
- `serial`: 序号（支持变量）
- `use_llm`: 是否使用LLM生成自然语言，默认 true
- `save_to`: 将生成的MCP提示词保存到变量名，默认 "mcp_prompt"

**自动保存的变量：**
- `{save_to}`: MCP提示词
- `{save_to}_nl`: 自然语言描述
- `{save_to}_json`: 提取的JSON数据

## 💡 变量使用

### 定义变量

```json
{
  "variables": {
    "username": "5130008",
    "password": "Uestc418",
    "project_code": "M112023ZHCG0006"
  }
}
```

### 使用变量

在步骤中使用 `${变量名}` 引用：

```json
{
  "type": "mcp",
  "prompt": "在用户名输入框中输入${username}\n在密码输入框中输入${password}"
}
```

### 自动变量

工作流引擎会自动设置以下变量：
- `last_session_id`: 最后一次 MCP 调用的会话ID
- `last_execution_id`: 最后一次 MCP 调用的执行ID

## 📚 示例工作流

### 示例 1: 从Excel生成提示词并执行（推荐）

```json
{
  "name": "Excel到MCP工作流",
  "variables": {
    "excel_path": "C:\\path\\to\\file.xlsx",
    "sheet_name": "3-报销",
    "serial": "1"
  },
  "steps": [
    {
      "type": "excel_to_prompt",
      "name": "从Excel生成MCP提示词",
      "excel_path": "${excel_path}",
      "sheet_name": "${sheet_name}",
      "serial": "${serial}",
      "use_llm": true,
      "save_to": "mcp_prompt"
    },
    {
      "type": "mcp",
      "name": "执行MCP提示词",
      "prompt": "${mcp_prompt}",
      "timeout": 300,
      "browser": "chrome",
      "headless": false
    }
  ]
}
```

### 示例 2: 批量处理多个序号

```json
{
  "name": "批量处理Excel",
  "variables": {
    "excel_path": "C:\\path\\to\\file.xlsx",
    "sheet_name": "3-报销",
    "serials": ["1", "2", "3"]
  },
  "steps": [
    {
      "type": "loop",
      "item_var": "current_serial",
      "items": "${serials}",
      "steps": [
        {
          "type": "excel_to_prompt",
          "excel_path": "${excel_path}",
          "sheet_name": "${sheet_name}",
          "serial": "${current_serial}",
          "save_to": "mcp_prompt"
        },
        {
          "type": "mcp",
          "prompt": "${mcp_prompt}",
          "timeout": 300,
          "continue_on_error": true
        },
        {
          "type": "wait",
          "seconds": 5
        }
      ]
    }
  ]
}
```

### 示例 3: 简单报销流程

```json
{
  "name": "报销流程",
  "variables": {
    "project_code": "M112023ZHCG0006",
    "amount": "100"
  },
  "steps": [
    {
      "type": "mcp",
      "name": "登录",
      "prompt": "1. 打开https://cwcx.uestc.edu.cn/WFManager/login.jsp\n2. 在用户名输入框中输入5130008\n3. 在密码输入框中输入Uestc418\n4. 等待 20 秒\n5. 点击id为zhLogin的登录按钮",
      "save_result_to": "login_result"
    },
    {
      "type": "condition",
      "condition": "${login_result.status} == 'success'",
      "if_true": [
        {
          "type": "mcp",
          "name": "填写报销信息",
          "prompt": "1. 点击网上预约报账按钮\n2. 点击申请报销单按钮\n3. 在报销项目号输入框中输入${project_code}\n4. 在附件张数输入框中输入2",
          "session_id": "${last_session_id}"
        }
      ]
    }
  ]
}
```

### 示例 2: 批量处理

```json
{
  "name": "批量处理",
  "variables": {
    "serials": ["1", "2", "3"]
  },
  "steps": [
    {
      "type": "loop",
      "item_var": "serial",
      "items": "${serials}",
      "steps": [
        {
          "type": "log",
          "message": "处理序号: ${serial}"
        },
        {
          "type": "mcp",
          "name": "处理单个报销",
          "prompt": "处理序号${serial}的报销...",
          "continue_on_error": true
        },
        {
          "type": "wait",
          "seconds": 5
        }
      ]
    }
  ]
}
```

## 🎯 在 Cursor 中使用

### 方式 1: 直接运行

1. 创建工作流 JSON 文件
2. 在 Cursor 终端运行：
   ```bash
   python workflow_engine.py workflows/my_workflow.json
   ```

### 方式 2: 集成到现有脚本

```python
from workflow_engine import WorkflowEngine

# 创建引擎
engine = WorkflowEngine("my_workflow.json")

# 执行工作流
result = engine.run()

# 检查结果
if result["status"] == "success":
    print("工作流执行成功")
else:
    print("工作流执行失败")
```

### 方式 3: 动态生成工作流

```python
import json
from workflow_engine import WorkflowEngine

# 动态创建工作流配置
workflow_config = {
    "name": "动态工作流",
    "variables": {
        "project_code": "M112023ZHCG0006"
    },
    "steps": [
        {
            "type": "mcp",
            "prompt": "打开页面并登录...",
            "timeout": 300
        }
    ]
}

# 保存为临时文件
with open("temp_workflow.json", "w") as f:
    json.dump(workflow_config, f, indent=2)

# 执行
engine = WorkflowEngine("temp_workflow.json")
result = engine.run()
```

## 🔍 调试和日志

工作流执行时会输出详细日志：

```
[2025-11-18 17:00:00] [INFO] 开始执行工作流: 报销流程
[2025-11-18 17:00:00] [INFO] 执行步骤: 登录
[2025-11-18 17:00:00] [INFO] 调用MCP: 1. 打开https://...
[2025-11-18 17:00:05] [INFO] MCP执行完成: success
```

执行结果会保存到 `workflow_result_YYYYMMDD_HHMMSS.json` 文件。

## ⚙️ 环境变量

可以通过环境变量配置：

```bash
# 设置MCP端点
export MCP_HTTP_ENDPOINT=http://localhost:3030/mcp/execute

# 运行工作流
python workflow_engine.py my_workflow.json
```

## 📦 依赖

确保已安装：

```bash
pip install requests pyyaml
```

## 🎨 最佳实践

1. **模块化**: 将复杂工作流拆分为多个小工作流
2. **变量化**: 使用变量避免硬编码
3. **错误处理**: 使用 `continue_on_error` 处理非关键步骤
4. **会话复用**: 使用 `session_id` 保持浏览器会话
5. **日志记录**: 使用 `log` 步骤记录关键信息

## 🐛 故障排查

### MCP 连接失败

- 检查 MCP 网关是否运行：`http://localhost:3030/health`
- 检查端点地址是否正确
- 查看网关日志

### 步骤执行失败

- 查看工作流执行日志
- 检查变量是否正确设置
- 验证 MCP 提示词格式

### 浏览器会话问题

- 使用 `session_id` 保持会话
- 检查浏览器是否被关闭
- 增加等待时间

## 📖 更多示例

查看 `workflows/` 目录下的示例文件：
- `simple_mcp_test.json` - 简单测试
- `reimburse_example.json` - 报销示例
- `batch_reimburse.json` - 批量处理

