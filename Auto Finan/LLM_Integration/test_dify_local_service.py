#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试 Dify 本地服务

用于验证本地服务是否正常运行
"""

import requests
import json
import sys

# 修复 Windows 控制台编码
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except:
        pass


def test_service():
    """测试本地服务"""
    base_url = "http://localhost:8001"
    
    print("=" * 80)
    print("Dify 本地服务测试")
    print("=" * 80)
    print(f"\n服务地址: {base_url}\n")
    
    # 1. 测试健康检查
    print("1. 测试健康检查端点...")
    try:
        response = requests.get(f"{base_url}/health", timeout=5)
        if response.status_code == 200:
            print(f"   ✅ 健康检查成功: {response.json()}")
        else:
            print(f"   ⚠️  健康检查返回状态码: {response.status_code}")
            return False
    except requests.exceptions.ConnectionError:
        print(f"   ❌ 无法连接到服务: {base_url}")
        print(f"   💡 请确保本地服务正在运行")
        print(f"   💡 运行命令: start_dify_local_service.bat")
        return False
    except Exception as e:
        print(f"   ❌ 健康检查失败: {e}")
        return False
    
    # 2. 测试 Excel 转提示词接口
    print("\n2. 测试 Excel 转提示词接口...")
    
    # 使用示例数据（需要根据实际情况修改）
    test_data = {
        "excel_path": r"C:\Users\FH\source\repos\Auto Finan\Auto Finan\420财务050823.xlsx",
        "sheet_name": "3-报销",
        "serial": "1"
    }
    
    print(f"   请求数据: {json.dumps(test_data, ensure_ascii=False, indent=2)}")
    
    try:
        response = requests.post(
            f"{base_url}/api/excel-to-prompt",
            json=test_data,
            timeout=30,
            headers={"Content-Type": "application/json"}
        )
        
        if response.status_code == 200:
            result = response.json()
            if result.get("success"):
                print(f"   ✅ 接口调用成功")
                print(f"   提示词长度: {result.get('prompt_length', 0)} 字符")
                print(f"   提示词预览（前 200 字符）:")
                prompt = result.get("mcp_prompt", "")
                print(f"   {prompt[:200]}...")
                return True
            else:
                print(f"   ⚠️  处理失败: {result.get('error', '未知错误')}")
                if result.get("suggestion"):
                    print(f"   💡 建议: {result.get('suggestion')}")
                if result.get("received"):
                    print(f"   📥 接收到的参数: {result.get('received')}")
                if result.get("debug"):
                    print(f"   🔍 调试信息: {result.get('debug')}")
                return False
        else:
            print(f"   ⚠️  接口返回状态码: {response.status_code}")
            try:
                error_detail = response.json()
                print(f"   错误详情: {json.dumps(error_detail, indent=2, ensure_ascii=False)}")
            except:
                print(f"   响应: {response.text[:500]}")
            return False
            
    except requests.exceptions.Timeout:
        print(f"   ⚠️  请求超时（可能文件较大，处理时间较长）")
        return False
    except requests.exceptions.ConnectionError:
        print(f"   ❌ 无法连接到服务")
        return False
    except Exception as e:
        print(f"   ❌ 测试失败: {e}")
        return False


if __name__ == "__main__":
    success = test_service()
    
    print("\n" + "=" * 80)
    if success:
        print("✅ 服务测试通过！可以正常使用")
    else:
        print("❌ 服务测试失败！")
        print("\n故障排查步骤：")
        print("1. 检查服务是否运行: start_dify_local_service.bat")
        print("2. 检查端口是否被占用: netstat -an | findstr :8001")
        print("3. 检查防火墙设置")
        print("4. 查看服务日志")
    print("=" * 80)
    
    sys.exit(0 if success else 1)

