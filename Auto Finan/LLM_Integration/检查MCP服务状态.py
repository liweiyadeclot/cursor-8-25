"""
检查 Playwright MCP HTTP 网关服务状态
"""

import requests
import sys

def check_service():
    """检查服务状态"""
    base_url = "http://localhost:3030"
    
    print("=" * 60)
    print("检查 Playwright MCP 服务状态")
    print("=" * 60)
    print()
    
    # 1. 检查根路径
    print("1. 检查服务根路径...")
    try:
        response = requests.get(f"{base_url}/", timeout=5)
        if response.status_code == 200:
            print(f"   ✅ 服务正在运行")
            try:
                data = response.json()
                print(f"   服务信息: {data}")
            except:
                print(f"   响应: {response.text[:100]}")
        else:
            print(f"   ⚠️  服务返回状态码: {response.status_code}")
    except requests.exceptions.ConnectionError:
        print(f"   ❌ 无法连接到服务")
        print(f"   请确保服务正在运行")
        print(f"   启动命令: start_mcp_gateway.bat")
        return False
    except Exception as e:
        print(f"   ❌ 检查失败: {e}")
        return False
    
    print()
    
    # 2. 检查健康检查端点（Python 网关）
    print("2. 检查健康检查端点（Python 网关）...")
    try:
        response = requests.get(f"{base_url}/health", timeout=5)
        if response.status_code == 200:
            print(f"   ✅ Python 网关健康检查通过")
            print(f"   响应: {response.json()}")
            is_python_gateway = True
        else:
            print(f"   ⚠️  健康检查返回: {response.status_code}")
            print(f"   （可能是官方服务器，没有 /health 端点）")
            is_python_gateway = False
    except Exception as e:
        print(f"   ⚠️  健康检查失败: {e}")
        is_python_gateway = False
    
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
            print(f"   💡 这是官方 Playwright MCP 服务器")
            print(f"   💡 它使用 MCP 协议，不是简单的 HTTP POST")
            print(f"   💡 对于 Dify，建议使用 Python 网关服务")
            print(f"   💡 启动命令: start_mcp_gateway.bat")
        else:
            print(f"   ⚠️  执行端点返回状态码: {response.status_code}")
            print(f"   响应: {response.text[:200]}")
    except Exception as e:
        print(f"   ⚠️  测试执行端点失败: {e}")
    
    print()
    print("=" * 60)
    print("检查完成")
    print("=" * 60)
    
    if not is_python_gateway:
        print()
        print("💡 建议：")
        print("   当前运行的是官方 Playwright MCP 服务器")
        print("   对于 Dify 工作流，建议使用 Python 网关服务")
        print("   1. 停止当前服务器（Ctrl+C）")
        print("   2. 运行: start_mcp_gateway.bat")
    
    return True

if __name__ == "__main__":
    success = check_service()
    sys.exit(0 if success else 1)

