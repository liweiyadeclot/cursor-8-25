"""
Dify 代码节点：处理 HTTP 响应（官方标准方式）
根据 Dify 官方文档，使用函数定义方式
"""

def main(http_response: str) -> dict:
    """
    处理 HTTP 响应
    
    参数：
        http_response: HTTP 响应内容（JSON 字符串）
    
    返回：
        dict: 包含处理结果的字典
    
    注意：根据实际响应格式访问，不要假设有 data 键
    """
    import json
    
    # 解析 JSON
    try:
        response = json.loads(http_response)
    except json.JSONDecodeError as e:
        return {
            "success": False,
            "error": f"JSON 解析失败: {str(e)}",
            "raw_response": http_response[:200] if len(http_response) > 200 else http_response
        }
    
    # 检查响应格式
    if not isinstance(response, dict):
        return {
            "success": False,
            "error": f"响应格式错误，期望字典，得到: {type(response).__name__}",
            "raw_response": str(response)[:200]
        }
    
    # 处理响应（根据实际响应格式）
    # 我们的服务返回格式：{"success": true, "mcp_prompt": "...", "prompt_length": 123}
    # 注意：不要访问 response['data']，因为我们的响应中没有 data 键
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
            "debug": response.get("debug", {})
        }

