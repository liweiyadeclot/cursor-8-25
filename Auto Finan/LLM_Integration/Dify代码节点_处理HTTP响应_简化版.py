"""
Dify 代码节点：处理 HTTP 响应（简化版）
最简化的版本，确保能正常工作
"""

def main(http_response: str) -> dict:
    import json
    
    # 解析 JSON
    response = json.loads(http_response)
    
    # 直接返回，确保所有输出变量都存在
    if response.get("success"):
        return {
            "success": True,
            "mcp_prompt": response.get("mcp_prompt", ""),
            "prompt_length": response.get("prompt_length", 0)
        }
    else:
        return {
            "success": False,
            "mcp_prompt": "",
            "prompt_length": 0,
            "error": response.get("error", "未知错误")
        }

