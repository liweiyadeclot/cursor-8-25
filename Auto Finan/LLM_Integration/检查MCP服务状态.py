"""
检查 Playwright MCP HTTP 网关服务状态
支持检查本地和远程服务
"""

import requests
import sys
import socket
import subprocess
import os

def check_port_open(host, port, timeout=3):
    """检查端口是否开放"""
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(timeout)
        result = sock.connect_ex((host, port))
        sock.close()
        return result == 0
    except Exception:
        return False

def get_local_ip():
    """获取本机 IP 地址"""
    try:
        # 连接到一个远程地址来获取本机 IP
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"

def check_service(base_url=None, host=None, port=3030):
    """检查服务状态"""
    if base_url is None:
        if host is None:
            host = "localhost"
        base_url = f"http://{host}:{port}"
    
    # 解析 host
    if host is None:
        if "://" in base_url:
            host = base_url.split("://")[1].split(":")[0].split("/")[0]
        else:
            host = "localhost"
    
    print("=" * 60)
    print("检查 Playwright MCP 服务状态")
    print("=" * 60)
    print(f"目标地址: {base_url}")
    print(f"主机: {host}, 端口: {port}")
    print()
    
    # 0. 检查端口是否开放
    print("0. 检查端口是否开放...")
    if check_port_open(host, port):
        print(f"   ✅ 端口 {port} 已开放")
    else:
        print(f"   ❌ 端口 {port} 未开放或无法连接")
        print(f"   💡 可能的原因：")
        print(f"      1. 服务未启动")
        print(f"      2. 防火墙阻止了连接")
        print(f"      3. 服务监听在 localhost 而不是 0.0.0.0")
        if host != "localhost" and host != "127.0.0.1":
            print(f"      4. 远程主机 {host} 不可达")
        print()
        print(f"   💡 解决方案：")
        print(f"      1. 确保服务正在运行: start_mcp_gateway.bat")
        print(f"      2. 检查服务是否监听在 0.0.0.0:3030（而不是 localhost）")
        print(f"      3. 检查 Windows 防火墙设置")
        if host != "localhost" and host != "127.0.0.1":
            local_ip = get_local_ip()
            print(f"      4. 如果从远程访问，确保服务监听在 0.0.0.0")
            print(f"         当前本机 IP: {local_ip}")
            print(f"         如果 {host} 不是本机，请检查网络连接")
        return False
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
    import argparse
    
    parser = argparse.ArgumentParser(description="检查 Playwright MCP HTTP 网关服务状态")
    parser.add_argument("--host", type=str, default=None, help="服务主机地址（默认: localhost）")
    parser.add_argument("--port", type=int, default=3030, help="服务端口（默认: 3030）")
    parser.add_argument("--url", type=str, default=None, help="完整服务 URL（例如: http://192.168.137.133:3030）")
    
    args = parser.parse_args()
    
    if args.url:
        success = check_service(base_url=args.url)
    else:
        success = check_service(host=args.host, port=args.port)
    
    sys.exit(0 if success else 1)

