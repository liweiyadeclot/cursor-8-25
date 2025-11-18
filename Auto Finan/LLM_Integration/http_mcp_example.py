#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Playwright MCP HTTP 调用示例脚本

功能：
1. 从 Excel 生成 MCP 提示词
2. 通过 HTTP 接口调用 Playwright MCP 执行
3. 处理响应结果

使用方法：
    python http_mcp_example.py
"""

import os
import sys
import json
import requests
from typing import Optional, Dict, Any

# 将当前脚本所在目录加入路径
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

try:
    from workflow_core import process_excel_to_mcp_direct, batch_process_excel_to_mcp_direct
except ImportError as e:
    print(f"❌ 导入失败: {e}")
    print("请确保 workflow_core.py 在同一目录下")
    sys.exit(1)


def check_endpoint_health(endpoint_base: str = None) -> bool:
    """
    检查端点健康状态
    
    Args:
        endpoint_base: 端点基础地址（不含路径），默认 http://localhost:3030
    
    Returns:
        是否健康
    """
    if not endpoint_base:
        endpoint_base = os.environ.get("MCP_HTTP_ENDPOINT", "http://localhost:3030")
        # 移除路径部分
        if "/mcp/execute" in endpoint_base:
            endpoint_base = endpoint_base.replace("/mcp/execute", "")
    
    health_url = f"{endpoint_base}/health"
    
    try:
        # 使用连接超时3秒，读取超时10秒
        response = requests.get(health_url, timeout=(3, 10))
        return response.status_code == 200
    except:
        return False


def call_playwright_mcp_http(prompt: str, endpoint: str = None, timeout: int = 30) -> Dict[str, Any]:
    """
    通过 HTTP 调用 Playwright MCP
    
    Args:
        prompt: MCP 提示词字符串
        endpoint: MCP HTTP 端点地址，默认从环境变量读取
        timeout: 请求超时时间（秒），默认 30 秒（缩短了，避免长时间等待）
    
    Returns:
        包含执行结果的字典
    """
    if not prompt or not prompt.strip():
        return {
            "status": "error",
            "message": "提示词为空"
        }
    
    # 获取端点地址
    if not endpoint:
        endpoint = os.environ.get("MCP_HTTP_ENDPOINT", "http://localhost:3030/mcp/execute")
    
    print(f"📡 调用 MCP 端点: {endpoint}")
    print(f"📝 提示词长度: {len(prompt)} 字符")
    
    # 先检查服务是否可用
    print("🔍 检查服务状态...")
    endpoint_base = endpoint.replace("/mcp/execute", "")
    if not check_endpoint_health(endpoint_base):
        print("⚠️  服务健康检查失败")
        return {
            "status": "error",
            "message": f"无法连接到 MCP 端点: {endpoint}",
            "error_type": "connection_error",
            "suggestion": "请确保 HTTP 网关服务正在运行。运行命令: start_mcp_gateway.bat"
        }
    
    print("✅ 服务可用，正在发送请求...")
    
    try:
        # 添加进度提示
        import time
        start_time = time.time()
        
        print(f"⏳ 等待响应（超时时间: {timeout}秒）...")
        
        response = requests.post(
            endpoint,
            json={"prompt": prompt},
            timeout=timeout,
            headers={"Content-Type": "application/json"}
        )
        
        elapsed_time = time.time() - start_time
        print(f"📥 收到响应（耗时: {elapsed_time:.2f}秒）")
        
        response.raise_for_status()
        
        result = response.json()
        print(f"✅ HTTP 请求成功 (状态码: {response.status_code})")
        return result
        
    except requests.exceptions.Timeout:
        return {
            "status": "error",
            "message": f"请求超时（{timeout}秒）。提示词可能过长或服务处理较慢。",
            "error_type": "timeout",
            "suggestion": "可以尝试增加超时时间或检查服务日志"
        }
    except requests.exceptions.ConnectionError as e:
        return {
            "status": "error",
            "message": f"无法连接到 MCP 端点: {endpoint}",
            "error_type": "connection_error",
            "error_details": str(e),
            "suggestion": "请确保 HTTP 网关服务正在运行。运行命令: start_mcp_gateway.bat"
        }
    except requests.exceptions.HTTPError as e:
        return {
            "status": "error",
            "message": f"HTTP 错误: {e}",
            "status_code": response.status_code if 'response' in locals() else None,
            "error_type": "http_error"
        }
    except Exception as e:
        return {
            "status": "error",
            "message": f"请求失败: {str(e)}",
            "error_type": "unknown",
            "error_details": str(e)
        }


def example_1_direct_prompt():
    """示例 1：直接使用提示词字符串调用"""
    print("\n" + "="*80)
    print("示例 1：直接使用提示词字符串调用")
    print("="*80)
    
    prompt = """请你调用Playwright MCP，执行以下命令，一次性执行完
打开https://cwcx.uestc.edu.cn/WFManager/login.jsp
业务大类：报销业务。以下是需要执行的页面操作：
在用户名输入框中输入5130008
在密码输入框中输入Uestc418
将验证码图片保存至C:\\Users\\FH\\source\\repos\\Auto Finan\\Auto Finan\\LLM_Integration目录下，命名为example.jpg
读取图片中的验证码信息
输入验证码
点击登录按钮"""
    
    result = call_playwright_mcp_http(prompt)
    
    print("\n📊 执行结果:")
    print(json.dumps(result, indent=2, ensure_ascii=False))
    
    return result


def example_2_from_excel_single():
    """示例 2：从 Excel 生成单个序号的提示词并调用"""
    print("\n" + "="*80)
    print("示例 2：从 Excel 生成单个序号的提示词并调用")
    print("="*80)
    
    # 配置 Excel 路径
    excel_path = os.path.join(os.path.dirname(CURRENT_DIR), "420财务050823.xlsx")
    sheet_name = "3-报销"
    serial = "1"
    
    if not os.path.exists(excel_path):
        print(f"❌ Excel 文件不存在: {excel_path}")
        return None
    
    print(f"📂 Excel 文件: {excel_path}")
    print(f"📋 工作表: {sheet_name}")
    print(f"🔢 序号: {serial}")
    
    # 生成 MCP 提示词
    print("\n🔄 正在生成 MCP 提示词...")
    mcp_prompt = process_excel_to_mcp_direct(excel_path, sheet_name, serial)
    
    if not mcp_prompt:
        print("❌ 未能生成 MCP 提示词")
        return None
    
    print(f"✅ MCP 提示词已生成（{len(mcp_prompt)} 字符）")
    print("\n📝 提示词预览（前 500 字符）:")
    print("-" * 80)
    print(mcp_prompt[:500] + "..." if len(mcp_prompt) > 500 else mcp_prompt)
    print("-" * 80)
    
    # 调用 MCP
    result = call_playwright_mcp_http(mcp_prompt)
    
    print("\n📊 执行结果:")
    print(json.dumps(result, indent=2, ensure_ascii=False))
    
    return result


def example_3_from_excel_batch():
    """示例 3：批量处理 Excel 所有序号"""
    print("\n" + "="*80)
    print("示例 3：批量处理 Excel 所有序号")
    print("="*80)
    
    excel_path = os.path.join(os.path.dirname(CURRENT_DIR), "420财务050823.xlsx")
    sheet_name = "3-报销"
    
    if not os.path.exists(excel_path):
        print(f"❌ Excel 文件不存在: {excel_path}")
        return []
    
    print(f"📂 Excel 文件: {excel_path}")
    print(f"📋 工作表: {sheet_name}")
    
    # 批量生成 MCP 提示词
    print("\n🔄 正在批量生成 MCP 提示词...")
    results = batch_process_excel_to_mcp_direct(excel_path, sheet_name)
    
    if not results:
        print("❌ 未能生成任何 MCP 提示词")
        return []
    
    print(f"✅ 共生成 {len(results)} 个 MCP 提示词")
    
    # 逐个调用 MCP（注意：实际使用时建议控制并发）
    execution_results = []
    for i, result in enumerate(results, 1):
        serial = result.get('serial', '未知')
        mcp_prompt = result.get('mcp_prompt', '')
        
        if not mcp_prompt:
            print(f"\n⚠️  序号 {serial} 无有效提示词，跳过")
            continue
        
        print(f"\n{'='*80}")
        print(f"处理序号 {serial} ({i}/{len(results)})")
        print(f"{'='*80}")
        
        exec_result = call_playwright_mcp_http(mcp_prompt)
        exec_result['serial'] = serial
        execution_results.append(exec_result)
        
        # 显示简要结果
        status = exec_result.get('status', 'unknown')
        if status == 'success':
            print(f"✅ 序号 {serial} 执行成功")
        else:
            print(f"❌ 序号 {serial} 执行失败: {exec_result.get('message', '未知错误')}")
    
    return execution_results


def example_4_from_file():
    """示例 4：从提示词文件读取并调用"""
    print("\n" + "="*80)
    print("示例 4：从提示词文件读取并调用")
    print("="*80)
    
    prompt_file = os.path.join(CURRENT_DIR, "mcp_prompts", "未预约-M112023ZHCG0006-100-20251010-08-20-08.txt")
    
    if not os.path.exists(prompt_file):
        print(f"❌ 提示词文件不存在: {prompt_file}")
        print("💡 提示：请先运行 excel_batch_processor.py 生成提示词文件")
        return None
    
    print(f"📂 读取提示词文件: {prompt_file}")
    
    try:
        with open(prompt_file, 'r', encoding='utf-8') as f:
            prompt = f.read()
        
        print(f"✅ 已读取提示词（{len(prompt)} 字符）")
        
        # 调用 MCP
        result = call_playwright_mcp_http(prompt)
        
        print("\n📊 执行结果:")
        print(json.dumps(result, indent=2, ensure_ascii=False))
        
        return result
        
    except Exception as e:
        print(f"❌ 读取文件失败: {e}")
        return None


def main():
    """主函数：运行所有示例"""
    print("="*80)
    print("Playwright MCP HTTP 调用示例")
    print("="*80)
    
    # 检查环境变量
    mcp_endpoint = os.environ.get("MCP_HTTP_ENDPOINT", "http://localhost:3030/mcp/execute")
    print(f"🌐 MCP 端点: {mcp_endpoint}")
    
    # 检查服务是否可用
    print("\n🔍 检查 HTTP 网关服务状态...")
    endpoint_base = mcp_endpoint.replace("/mcp/execute", "")
    if check_endpoint_health(endpoint_base):
        print("✅ HTTP 网关服务正在运行")
    else:
        print("❌ HTTP 网关服务未运行")
        print("\n💡 请先启动 HTTP 网关服务：")
        print("   1. 运行: start_mcp_gateway.bat")
        print("   2. 或运行: python playwright_mcp_http_gateway.py")
        print("\n⚠️  如果服务未启动，程序将无法连接")
        response = input("\n是否继续？(y/n): ").strip().lower()
        if response != 'y':
            print("已取消")
            return
    
    print("\n" + "="*80)
    print("请选择要运行的示例：")
    print("1. 直接使用提示词字符串调用")
    print("2. 从 Excel 生成单个序号的提示词并调用")
    print("3. 批量处理 Excel 所有序号")
    print("4. 从提示词文件读取并调用")
    print("0. 退出")
    print("="*80)
    
    try:
        choice = input("\n请输入选项 (0-4): ").strip()
        
        if choice == "1":
            example_1_direct_prompt()
        elif choice == "2":
            example_2_from_excel_single()
        elif choice == "3":
            example_3_from_excel_batch()
        elif choice == "4":
            example_4_from_file()
        elif choice == "0":
            print("👋 退出")
            return
        else:
            print("❌ 无效选项")
            return
            
    except KeyboardInterrupt:
        print("\n\n⚠️  用户中断")
    except Exception as e:
        print(f"\n❌ 发生错误: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()

