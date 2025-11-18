# Dify 工作流配置指南：Excel → 提示词 → Playwright MCP

## 📋 概述

本指南将帮助你配置 Dify 工作流，实现：
1. 从 Excel 文件读取数据
2. 转换为 MCP 提示词
3. 通过 HTTP 调用 Playwright MCP 执行

---

## 🏗️ 工作流架构

```
┌─────────────┐     ┌──────────────┐     ┌─────────────┐     ┌──────────────┐
│  开始节点   │ --> │  Excel读取   │ --> │  提示词生成 │ --> │  HTTP调用    │
│             │     │  节点        │     │  节点        │     │  Playwright  │
└─────────────┘     └──────────────┘     └─────────────┘     └──────────────┘
                                                                    │
                                                                    ▼
                                                          ┌──────────────┐
                                                          │  结果处理    │
                                                          │  节点        │
                                                          └──────────────┘
```

---

## 📝 详细配置步骤

### 步骤 1：准备 Excel 文件

确保 Excel 文件包含以下列：
- `序号` - 用于分组数据
- `账号` / `登录界面工号` - 登录用户名
- `密码` / `登录界面密码` - 登录密码
- `业务大类` / `选择业务大类` - 业务类型（报销业务、业务出差旅费、酬金业务）
- 其他业务相关字段

**示例文件路径**：
```
C:\Users\FH\source\repos\Auto Finan\Auto Finan\420财务050823.xlsx
```

---

### 步骤 2：创建 Dify 工作流

#### 2.1 添加开始节点

1. 在 Dify 中创建新工作流
2. 添加 **开始节点**
3. 配置输入变量：
   - `excel_path` (字符串) - Excel 文件路径
   - `sheet_name` (字符串) - 工作表名称，如 "3-报销"
   - `serial` (字符串) - 序号，如 "1"

---

#### 2.2 添加代码节点：Excel 读取和提示词生成

**节点类型**：`代码`

**代码内容**：

```python
import sys
import os

# 添加 Python 路径（如果需要）
sys.path.insert(0, r'C:\Users\FH\source\repos\Auto Finan\Auto Finan\LLM_Integration')

try:
    from workflow_core import process_excel_to_mcp_direct
except ImportError:
    # 如果导入失败，输出错误
    output = {
        "success": False,
        "error": "无法导入 workflow_core 模块",
        "suggestion": "请确保 workflow_core.py 在正确路径"
    }

# 从上游节点获取参数
# 注意：在 Dify 代码节点中，使用 inputs 字典访问变量
excel_path = inputs.get('excel_path', '')
sheet_name = inputs.get('sheet_name', '')
serial = inputs.get('serial', '')

# 验证参数
if not excel_path or not sheet_name or not serial:
    output = {
        "success": False,
        "error": "缺少必要参数",
        "required": ["excel_path", "sheet_name", "serial"],
        "received": {
            "excel_path": excel_path,
            "sheet_name": sheet_name,
            "serial": serial
        }
    }
elif not os.path.exists(excel_path):
    # 检查文件是否存在
    output = {
        "success": False,
        "error": f"Excel 文件不存在: {excel_path}"
    }
else:
    try:
        # 生成 MCP 提示词
        mcp_prompt = process_excel_to_mcp_direct(
            excel_path=excel_path,
            sheet_name=sheet_name,
            serial=serial
        )
        
        if not mcp_prompt or not mcp_prompt.strip():
            output = {
                "success": False,
                "error": "未能生成有效的 MCP 提示词",
                "suggestion": "请检查 Excel 数据和序号是否正确"
            }
        else:
            # 输出结果
            output = {
                "success": True,
                "mcp_prompt": mcp_prompt,
                "prompt_length": len(mcp_prompt),
                "excel_path": excel_path,
                "sheet_name": sheet_name,
                "serial": serial
            }
        
    except Exception as e:
        output = {
            "success": False,
            "error": f"生成提示词失败: {str(e)}",
            "error_type": type(e).__name__
        }
```

**输出变量**：
- `success` (布尔) - 是否成功
- `mcp_prompt` (字符串) - 生成的 MCP 提示词
- `prompt_length` (数字) - 提示词长度
- `error` (字符串) - 错误信息（如果失败）

---

#### 2.3 添加条件判断节点

**节点类型**：`条件判断`

**条件**：
```
{{#workflow.success#}} == true
```

**分支**：
- **True 分支**：继续执行 HTTP 调用
- **False 分支**：输出错误信息并结束

---

#### 2.4 添加 HTTP 请求节点：调用 Playwright MCP

**节点类型**：`HTTP 请求`

**配置**：

- **请求方法**：`POST`
- **URL**：`http://localhost:3030/mcp/execute`
  - 如果网关在其他机器，改为对应 IP
  - 例如：`http://192.168.1.100:3030/mcp/execute`

- **请求头**：
```json
{
  "Content-Type": "application/json"
}
```

- **请求体**：
```json
{
  "prompt": "{{#workflow.mcp_prompt#}}",
  "timeout": 300,
  "headless": false
}
```

**注意**：在 HTTP 请求节点中，可以使用 `{{#workflow.variable_name#}}` 语法引用变量。

- **超时设置**：`300` 秒（5分钟）

**输出变量**：
- `status` - 执行状态
- `message` - 执行消息
- `logs` - 执行日志
- `execution_id` - 执行 ID

---

#### 2.5 添加代码节点：处理执行结果

**节点类型**：`代码`

**代码内容**：

```python
# 从 HTTP 请求节点获取响应
# 注意：在代码节点中，使用 inputs 字典访问变量
http_response = inputs.get('http_response', {})

# 检查响应格式
if isinstance(http_response, str):
    import json
    try:
        http_response = json.loads(http_response)
    except:
        output = {
            "success": False,
            "error": "HTTP 响应格式错误",
            "raw_response": http_response[:500] if len(str(http_response)) > 500 else http_response
        }
    else:
        # 检查执行状态
        status = http_response.get("status", "unknown")
        message = http_response.get("message", "")
        logs = http_response.get("logs", [])
        execution_id = http_response.get("execution_id", "")
        
        if status == "success":
            output = {
                "success": True,
                "status": status,
                "message": message,
                "execution_id": execution_id,
                "logs": logs,
                "logs_count": len(logs),
                "note": "执行成功"
            }
        elif status == "partial":
            output = {
                "success": True,
                "status": status,
                "message": message,
                "execution_id": execution_id,
                "logs": logs,
                "warning": "部分步骤执行成功",
                "note": "请检查失败步骤"
            }
        else:
            # 执行失败
            error_details = http_response.get("error_details", {})
            output = {
                "success": False,
                "status": status,
                "message": message,
                "error": error_details,
                "logs": logs,
                "note": "执行失败，请检查错误信息"
            }
else:
    # 如果 http_response 已经是字典
    status = http_response.get("status", "unknown")
    message = http_response.get("message", "")
    logs = http_response.get("logs", [])
    execution_id = http_response.get("execution_id", "")
    
    if status == "success":
        output = {
            "success": True,
            "status": status,
            "message": message,
            "execution_id": execution_id,
            "logs": logs,
            "logs_count": len(logs),
            "note": "执行成功"
        }
    elif status == "partial":
        output = {
            "success": True,
            "status": status,
            "message": message,
            "execution_id": execution_id,
            "logs": logs,
            "warning": "部分步骤执行成功",
            "note": "请检查失败步骤"
        }
    else:
        error_details = http_response.get("error_details", {})
        output = {
            "success": False,
            "status": status,
            "message": message,
            "error": error_details,
            "logs": logs,
            "note": "执行失败，请检查错误信息"
        }
```

**输出变量**：
- `success` - 是否成功
- `status` - 执行状态
- `message` - 消息
- `logs` - 日志列表
- `error` - 错误信息（如果失败）

---

#### 2.6 添加结束节点

**节点类型**：`结束`

**输出**：
- 显示执行结果
- 记录日志
- 返回给调用方

---

## 🔗 节点连接顺序

```
开始节点
  │
  ├─> [代码节点：Excel读取和提示词生成]
  │     │
  │     └─> [条件判断节点]
  │           │
  │           ├─> [True] ──> [HTTP请求节点：调用Playwright MCP]
  │           │                │
  │           │                └─> [代码节点：处理执行结果]
  │           │                      │
  │           │                      └─> [结束节点]
  │           │
  │           └─> [False] ──> [结束节点（错误）]
```

---

## 📋 完整工作流配置示例

### 工作流输入变量

```json
{
  "excel_path": "C:\\Users\\FH\\source\\repos\\Auto Finan\\Auto Finan\\420财务050823.xlsx",
  "sheet_name": "3-报销",
  "serial": "1"
}
```

### 工作流输出

```json
{
  "success": true,
  "status": "success",
  "message": "提示词已解析，共识别到 32 个操作步骤",
  "execution_id": "exec_20251118_100839",
  "logs": [
    "步骤 1: 打开页面 https://cwcx.uestc.edu.cn/WFManager/login.jsp",
    "步骤 2: 在用户名输入框中输入5130008",
    ...
  ],
  "logs_count": 32
}
```

---

## 🚀 使用方式

### 方式 1：在 Dify 界面中运行

1. 打开工作流
2. 点击"运行"
3. 输入参数：
   - Excel 文件路径
   - 工作表名称
   - 序号
4. 点击"执行"
5. 查看执行结果

### 方式 2：通过 API 调用

```python
import requests

# Dify API 端点
dify_api_url = "https://api.dify.ai/v1/workflows/run"

# 工作流 ID
workflow_id = "your-workflow-id"

# API Key
api_key = "your-api-key"

# 请求参数
payload = {
    "inputs": {
        "excel_path": "C:\\Users\\FH\\source\\repos\\Auto Finan\\Auto Finan\\420财务050823.xlsx",
        "sheet_name": "3-报销",
        "serial": "1"
    }
}

# 发送请求
response = requests.post(
    f"{dify_api_url}/{workflow_id}",
    json=payload,
    headers={
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
)

result = response.json()
print(result)
```

---

## ⚙️ 环境配置

### 1. 确保 HTTP 网关服务运行

在运行 Dify 工作流前，确保 HTTP 网关服务已启动：

```bash
cd "Auto Finan\LLM_Integration"
start_mcp_gateway.bat
```

### 2. 配置 Python 环境

如果 Dify 运行在服务器上，需要：
- 安装 Python 依赖：`pip install openpyxl requests`
- 确保 `workflow_core.py` 在可访问路径
- 或者将代码直接写在 Dify 代码节点中

### 3. 网络配置

- 确保 Dify 服务器可以访问 HTTP 网关（`http://localhost:3030`）
- 如果 Dify 在远程服务器，需要：
  - 将网关绑定到 `0.0.0.0`（已配置）
  - 配置防火墙允许 3030 端口
  - 使用服务器 IP 地址

---

## 🔧 高级配置

### 批量处理多个序号

如果需要批量处理，可以添加循环节点：

```python
# 在代码节点中
from workflow_core import batch_process_excel_to_mcp_direct

excel_path = {{#workflow.excel_path#}}
sheet_name = {{#workflow.sheet_name#}}

# 批量生成所有序号的提示词
results = batch_process_excel_to_mcp_direct(excel_path, sheet_name)

# 返回结果列表
return {
    "success": True,
    "results": results,
    "count": len(results)
}
```

然后在 HTTP 请求节点中使用循环处理每个结果。

### 错误处理和重试

在 HTTP 请求节点后添加：
- **条件判断**：检查执行状态
- **重试逻辑**：如果失败，重试指定次数
- **错误通知**：发送错误通知（邮件、消息等）

---

## 📝 注意事项

1. **文件路径**：
   - Windows 路径使用双反斜杠：`C:\\Users\\...`
   - 或使用原始字符串：`r"C:\Users\..."`

2. **Python 模块导入**：
   - 如果 Dify 运行在服务器，需要确保模块路径正确
   - 或者将代码直接写在节点中

3. **HTTP 网关地址**：
   - 本地：`http://localhost:3030`
   - 远程：`http://服务器IP:3030`

4. **超时设置**：
   - Excel 处理：通常很快（< 5秒）
   - HTTP 调用：根据提示词长度，建议 300 秒

5. **并发控制**：
   - 避免同时执行多个自动化任务
   - 可能导致浏览器冲突

---

## 🐛 故障排查

### 问题 1：无法导入 workflow_core

**解决**：
- 检查 Python 路径配置
- 或将代码直接写在节点中

### 问题 2：HTTP 请求失败

**检查**：
- HTTP 网关服务是否运行
- 网络连接是否正常
- 防火墙设置

### 问题 3：提示词生成失败

**检查**：
- Excel 文件路径是否正确
- 工作表名称是否正确
- 序号是否存在

---

## 📚 相关文件

- `workflow_core.py` - 核心工作流程
- `excel_batch_processor.py` - Excel 批量处理
- `playwright_mcp_http_gateway.py` - HTTP 网关服务
- `http_mcp_example.py` - 示例代码

---

## 💡 下一步

1. ✅ 按照本指南配置 Dify 工作流
2. ✅ 测试单个序号的执行
3. ✅ 扩展到批量处理
4. ✅ 添加错误处理和通知

如有问题，请参考相关文档或检查日志。

