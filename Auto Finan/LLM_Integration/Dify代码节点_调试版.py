"""
Dify 代码节点：调试版本
用于查看实际输入和输出
"""

def main(http_response: str) -> dict:
    import json
    
    # 先返回调试信息，查看实际输入
    try:
        response = json.loads(http_response)
        
        # 返回调试信息
        return {
            "debug_input_type": type(http_response).__name__,
            "debug_input_length": len(http_response),
            "debug_response_type": type(response).__name__,
            "debug_response_keys": list(response.keys()) if isinstance(response, dict) else [],
            "debug_success": response.get("success") if isinstance(response, dict) else None,
            "success": True,
            "mcp_prompt": response.get("mcp_prompt", "") if isinstance(response, dict) else "",
            "prompt_length": response.get("prompt_length", 0) if isinstance(response, dict) else 0
        }
    except Exception as e:
        return {
            "debug_error": str(e),
            "debug_input_preview": str(http_response)[:200],
            "success": False,
            "mcp_prompt": "",
            "prompt_length": 0,
            "error": f"处理失败: {str(e)}"
        }

