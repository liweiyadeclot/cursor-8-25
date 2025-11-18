"""
Dify 代码节点：处理 HTTP 响应（纯 output 变量版本）
不使用函数定义，直接使用 output 变量
"""

import json

# 获取输入变量 http_response
# 注意：根据你的 Dify 版本，可能需要使用不同的方式获取

# 尝试多种方式获取 http_response
http_response = None

# 方式 1：从函数参数（如果使用函数定义）
try:
    if 'http_response' in locals():
        http_response = locals()['http_response']
except:
    pass

# 方式 2：从全局变量
if not http_response:
    try:
        if 'http_response' in globals():
            http_response = globals()['http_response']
    except:
        pass

# 方式 3：从 inputs（某些版本）
if not http_response:
    try:
        if 'inputs' in globals():
            http_response = globals()['inputs'].get('http_response', '')
    except:
        pass

# 解析和处理
if http_response:
    try:
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
    except Exception as e:
        output = {
            "success": False,
            "mcp_prompt": "",
            "prompt_length": 0,
            "error": f"处理失败: {str(e)}"
        }
else:
    # 如果无法获取输入，返回错误
    output = {
        "success": False,
        "mcp_prompt": "",
        "prompt_length": 0,
        "error": "无法获取 http_response 输入变量"
    }

