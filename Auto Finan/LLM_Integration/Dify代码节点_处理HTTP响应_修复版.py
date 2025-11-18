"""
Dify 代码节点：处理 HTTP 响应（修复 inputs 未定义问题）
直接复制到 Dify 代码节点中使用
"""

import json

# ============================================================================
# 安全获取 inputs（兼容不同 Dify 版本）
# ============================================================================

def get_inputs_safe():
    """安全获取 inputs，兼容不同 Dify 版本"""
    # 方式 1：从 globals 获取
    try:
        if 'inputs' in globals():
            return globals()['inputs']
    except:
        pass
    
    # 方式 2：从 locals 获取
    try:
        if 'inputs' in locals():
            return locals()['inputs']
    except:
        pass
    
    # 方式 3：尝试从上下文获取
    try:
        import sys
        frame = sys._getframe(1)
        if 'inputs' in frame.f_locals:
            return frame.f_locals['inputs']
    except:
        pass
    
    # 方式 4：返回空字典
    return {}

# 获取 inputs
inputs = get_inputs_safe()

# ============================================================================
# 获取 HTTP 响应
# ============================================================================

# 尝试多种可能的变量名
http_response = None

# 方式 1：从 inputs 字典
if inputs:
    http_response = inputs.get('http_response', '')
    if not http_response:
        # 尝试其他可能的变量名
        for key in ['response', 'result', 'data', 'http_result']:
            if key in inputs:
                http_response = inputs[key]
                break

# 方式 2：从全局变量
if not http_response:
    try:
        http_response = globals().get('http_response', '')
    except:
        pass

# ============================================================================
# 解析 JSON 响应
# ============================================================================

if not http_response:
    output = {
        "success": False,
        "error": "未找到 HTTP 响应",
        "debug": {
            "inputs_keys": list(inputs.keys()) if inputs else [],
            "globals_keys": [k for k in globals().keys() if not k.startswith('_')][:20]
        }
    }
else:
    # 解析 JSON（如果是字符串）
    if isinstance(http_response, str):
        try:
            response = json.loads(http_response)
        except json.JSONDecodeError as e:
            output = {
                "success": False,
                "error": f"JSON 解析失败: {str(e)}",
                "raw_response": http_response[:200] if len(http_response) > 200 else http_response
            }
        else:
            # 成功解析
            if response.get("success"):
                output = {
                    "success": True,
                    "mcp_prompt": response.get("mcp_prompt", ""),
                    "prompt_length": response.get("prompt_length", 0),
                    "message": response.get("message", "提示词生成成功")
                }
            else:
                output = {
                    "success": False,
                    "error": response.get("error", "未知错误"),
                    "debug": response.get("debug", {})
                }
    else:
        # 已经是字典
        response = http_response if isinstance(http_response, dict) else {}
        if response.get("success"):
            output = {
                "success": True,
                "mcp_prompt": response.get("mcp_prompt", ""),
                "prompt_length": response.get("prompt_length", 0)
            }
        else:
            output = {
                "success": False,
                "error": response.get("error", "未知错误")
            }

