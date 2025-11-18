"""
Dify 代码节点：处理 HTTP 响应（最终版）
确保正确返回输出变量
"""

def main(http_response: str) -> dict:
    """
    处理 HTTP 响应
    
    参数：
        http_response: HTTP 响应内容（JSON 字符串）
    
    返回：
        dict: 包含处理结果的字典
    """
    import json
    
    # 解析 JSON
    try:
        response = json.loads(http_response)
    except json.JSONDecodeError as e:
        return {
            "success": False,
            "error": f"JSON 解析失败: {str(e)}",
            "mcp_prompt": "",  # 确保所有输出变量都存在
            "prompt_length": 0
        }
    
    # 检查响应格式
    if not isinstance(response, dict):
        return {
            "success": False,
            "error": f"响应格式错误，期望字典，得到: {type(response).__name__}",
            "mcp_prompt": "",
            "prompt_length": 0
        }
    
    # 处理响应（根据实际响应格式）
    if response.get("success"):
        # 成功：提取 MCP 提示词
        return {
            "success": True,
            "mcp_prompt": response.get("mcp_prompt", ""),
            "prompt_length": response.get("prompt_length", 0),
            "message": response.get("message", "提示词生成成功")
        }
    else:
        # 失败：提取错误信息
        return {
            "success": False,
            "error": response.get("error", "未知错误"),
            "mcp_prompt": "",  # 确保所有输出变量都存在
            "prompt_length": 0,
            "debug": response.get("debug", {})
        }

