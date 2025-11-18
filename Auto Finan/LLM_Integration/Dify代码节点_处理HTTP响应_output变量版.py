"""
Dify 代码节点：处理 HTTP 响应（使用 output 变量）
某些 Dify 版本需要使用 output 变量，而不是 return
"""

import json

# 获取输入变量（从函数参数或直接使用）
# 如果使用函数定义，参数会自动传入
# 如果不使用函数，直接从 http_response 获取

# 方式 1：如果使用函数定义
def main(http_response: str) -> dict:
    response = json.loads(http_response)
    
    if response.get("success"):
        output = {
            "success": True,
            "mcp_prompt": response.get("mcp_prompt", ""),
            "prompt_length": response.get("prompt_length", 0)
        }
    else:
        output = {
            "success": False,
            "mcp_prompt": "",
            "prompt_length": 0,
            "error": response.get("error", "未知错误")
        }
    
    return output

# 方式 2：如果不使用函数定义（某些版本）
# 直接使用 output 变量
try:
    # 尝试从函数参数获取
    if 'http_response' in locals():
        http_response = locals()['http_response']
    elif 'http_response' in globals():
        http_response = globals()['http_response']
    
    # 解析 JSON
    response = json.loads(http_response)
    
    # 设置 output
    if response.get("success"):
        output = {
            "success": True,
            "mcp_prompt": response.get("mcp_prompt", ""),
            "prompt_length": response.get("prompt_length", 0)
        }
    else:
        output = {
            "success": False,
            "mcp_prompt": "",
            "prompt_length": 0,
            "error": response.get("error", "未知错误")
        }
except Exception as e:
    output = {
        "success": False,
        "mcp_prompt": "",
        "prompt_length": 0,
        "error": str(e)
    }

