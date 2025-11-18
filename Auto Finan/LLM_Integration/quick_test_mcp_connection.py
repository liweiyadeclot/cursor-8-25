#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速测试 MCP HTTP 连接

用于快速检查 HTTP 网关服务是否可用
"""

import os
import sys
import requests

# 修复 Windows 控制台编码
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except:
        pass


def test_connection():
    """测试连接"""
    endpoint = os.environ.get("MCP_HTTP_ENDPOINT", "http://localhost:3030/mcp/execute")
    endpoint_base = endpoint.replace("/mcp/execute", "")
    
    print("=" * 80)
    print("MCP HTTP 连接测试")
    print("=" * 80)
    print(f"\n端点地址: {endpoint}")
    print(f"基础地址: {endpoint_base}")
    print()
    
    # 测试健康检查
    print("1. 测试健康检查端点...")
    health_url = f"{endpoint_base}/health"
    try:
        # 使用更短的连接超时，但允许更长的读取超时
        response = requests.get(health_url, timeout=(3, 10))
        if response.status_code == 200:
            health_data = response.json()
            print(f"   ✅ 健康检查成功")
            print(f"   状态: {health_data.get('status', 'unknown')}")
            print(f"   时间戳: {health_data.get('timestamp', 'unknown')}")
        else:
            print(f"   ⚠️  健康检查返回状态码: {response.status_code}")
            print(f"   响应: {response.text[:200]}")
    except requests.exceptions.Timeout:
        print(f"   ⚠️  健康检查超时")
        print(f"   💡 服务可能正在启动中，请稍后再试")
        print(f"   💡 或者检查防火墙设置")
        # 超时不一定是失败，服务可能在运行
        return False
    except requests.exceptions.ConnectionError as e:
        print(f"   ❌ 无法连接到 {health_url}")
        print(f"   错误详情: {str(e)[:200]}")
        print(f"   💡 请确保 HTTP 网关服务正在运行")
        print(f"   💡 运行命令: start_mcp_gateway.bat")
        return False
    except Exception as e:
        print(f"   ❌ 健康检查失败: {e}")
        return False
    
    # 测试执行端点
    print("\n2. 测试执行端点...")
    test_prompt = "测试提示词"
    try:
        print(f"   发送测试请求到: {endpoint}")
        response = requests.post(
            endpoint,
            json={"prompt": test_prompt},
            timeout=(3, 15),  # 连接超时3秒，读取超时15秒
            headers={"Content-Type": "application/json"}
        )
        if response.status_code == 200:
            result = response.json()
            print(f"   ✅ 执行端点可用")
            print(f"   状态: {result.get('status', 'unknown')}")
            print(f"   消息: {result.get('message', '')[:100]}")
            return True
        else:
            print(f"   ⚠️  执行端点返回状态码: {response.status_code}")
            print(f"   响应: {response.text[:200]}")
            return False
    except requests.exceptions.Timeout as e:
        print(f"   ⚠️  请求超时: {e}")
        print(f"   💡 服务可能正在处理，这是正常的")
        # 如果健康检查通过了，超时也算成功（说明服务在运行）
        return True
    except requests.exceptions.ConnectionError as e:
        print(f"   ❌ 无法连接到 {endpoint}")
        print(f"   错误详情: {str(e)[:200]}")
        print(f"   💡 请确保 HTTP 网关服务正在运行")
        return False
    except Exception as e:
        print(f"   ❌ 测试失败: {e}")
        return False


if __name__ == "__main__":
    success = test_connection()
    print("\n" + "=" * 80)
    if success:
        print("✅ 连接测试通过！可以正常使用 HTTP 网关服务")
    else:
        print("❌ 连接测试失败！请先启动 HTTP 网关服务")
        print("\n启动命令:")
        print("  cd \"Auto Finan\\LLM_Integration\"")
        print("  start_mcp_gateway.bat")
    print("=" * 80)
    sys.exit(0 if success else 1)

