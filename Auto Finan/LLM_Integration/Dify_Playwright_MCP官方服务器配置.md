# Dify 使用 Playwright MCP 官方 HTTP 服务器配置

## 🔍 当前情况

你启动的是 **Playwright MCP 官方 HTTP 服务器**（通过 `npx @playwright/mcp@0.0.46`），而不是我们的 Python 网关。

官方服务器的 API 格式不同！

---

## ✅ 解决方案

### 方案 1：使用官方服务器的正确端点（推荐）

官方服务器使用 **MCP 协议**，不是简单的 HTTP POST。

**配置信息**（从服务器输出）：
```
Put this in your client config:
{
  "mcpServers": {
    "playwright": {
      "url": "http://localhost:3030/mcp"
    }
  }
}
```

**问题**：Dify 的 HTTP 请求节点无法直接使用 MCP 协议。

---

### 方案 2：使用我们的 Python 网关服务（推荐用于 Dify）

我们的 Python 网关提供了简单的 HTTP POST 接口，更适合 Dify。

#### 步骤 1：停止官方服务器

按 `Ctrl+C` 停止当前运行的官方服务器。

#### 步骤 2：启动 Python 网关

```bash
cd "C:\Users\FH\source\repos\Auto Finan\Auto Finan\LLM_Integration"
start_mcp_gateway.bat
```

或者：

```bash
python playwright_mcp_http_gateway.py
```

#### 步骤 3：验证服务

```bash
python 检查MCP服务状态.py
```

应该看到：
- ✅ 健康检查通过
- ✅ 执行端点正常

---

## 📝 Dify 配置

### 使用 Python 网关（推荐）

**HTTP 请求节点配置**：

| 项目 | 值 |
|------|-----|
| **方法** | `POST` |
| **URL** | `http://localhost:3030/mcp/execute`<br>（如果 Dify 在远程：`http://你的IP:3030/mcp/execute`） |
| **请求头** | `{"Content-Type": "application/json"}` |
| **请求体** | `{"prompt": "{{#workflow.mcp_prompt#}}"}` |

---

## 🔧 修复检查脚本

更新检查脚本以适配两种服务：

```python
import requests
import sys

def check_service():
    base_url = "http://localhost:3030"
    
    print("=" * 60)
    print("检查 Playwright MCP 服务状态")
    print("=" * 60)
    print()
    
    # 1. 检查根路径（两种服务都支持）
    print("1. 检查服务根路径...")
    try:
        response = requests.get(f"{base_url}/", timeout=5)
        if response.status_code == 200:
            print(f"   ✅ 服务正在运行")
            data = response.json()
            print(f"   服务信息: {data}")
        else:
            print(f"   ⚠️  服务返回状态码: {response.status_code}")
    except Exception as e:
        print(f"   ❌ 无法连接到服务: {e}")
        return False
    
    print()
    
    # 2. 检查健康检查端点（Python 网关）
    print("2. 检查健康检查端点（Python 网关）...")
    try:
        response = requests.get(f"{base_url}/health", timeout=5)
        if response.status_code == 200:
            print(f"   ✅ Python 网关健康检查通过")
            print(f"   响应: {response.json()}")
        else:
            print(f"   ⚠️  健康检查返回: {response.status_code}")
            print(f"   （可能是官方服务器，没有 /health 端点）")
    except Exception as e:
        print(f"   ⚠️  健康检查失败: {e}")
    
    print()
    
    # 3. 测试执行端点（Python 网关格式）
    print("3. 测试执行端点（Python 网关格式）...")
    try:
        test_data = {
            "prompt": "测试提示词"
        }
        response = requests.post(
            f"{base_url}/mcp/execute",
            json=test_data,
            timeout=10,
            headers={"Content-Type": "application/json"}
        )
        if response.status_code == 200:
            print(f"   ✅ 执行端点正常（Python 网关）")
            result = response.json()
            print(f"   状态: {result.get('status', 'unknown')}")
        elif response.status_code == 406:
            print(f"   ⚠️  执行端点返回 406（Not Acceptable）")
            print(f"   可能是官方服务器，需要 MCP 协议格式")
        else:
            print(f"   ⚠️  执行端点返回状态码: {response.status_code}")
            print(f"   响应: {response.text[:200]}")
    except Exception as e:
        print(f"   ⚠️  测试执行端点失败: {e}")
    
    print()
    print("=" * 60)
    print("检查完成")
    print("=" * 60)
    print()
    print("💡 建议：")
    print("   如果看到 406 错误，说明是官方服务器")
    print("   对于 Dify，建议使用 Python 网关服务")
    print("   启动命令: start_mcp_gateway.bat")
    
    return True

if __name__ == "__main__":
    check_service()
```

---

## 🎯 推荐方案

**对于 Dify 工作流，使用 Python 网关服务**：

1. ✅ 提供简单的 HTTP POST 接口
2. ✅ 与 Dify 兼容
3. ✅ 易于调试
4. ✅ 可以解析和验证提示词

**启动命令**：
```bash
start_mcp_gateway.bat
```

---

## 📚 相关文件

- `playwright_mcp_http_gateway.py` - Python 网关服务
- `start_mcp_gateway.bat` - 启动 Python 网关
- `start_playwright_mcp_http.bat` - 启动官方服务器（不推荐用于 Dify）

---

## ✅ 总结

**当前问题**：
- 你启动的是官方 Playwright MCP 服务器
- 它使用 MCP 协议，不是简单的 HTTP POST
- Dify 无法直接使用

**解决方案**：
- 停止官方服务器
- 启动 Python 网关服务：`start_mcp_gateway.bat`
- 使用 `/mcp/execute` 端点

