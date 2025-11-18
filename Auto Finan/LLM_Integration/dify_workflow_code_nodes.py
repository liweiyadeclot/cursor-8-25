#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dify 工作流代码节点示例

这些代码可以直接复制到 Dify 的代码节点中使用
"""

# ============================================================================
# 节点 1：Excel 读取和提示词生成
# ============================================================================

def node_excel_to_prompt():
    """
    Dify 代码节点：从 Excel 生成 MCP 提示词
    
    输入变量：
    - excel_path: Excel 文件路径
    - sheet_name: 工作表名称
    - serial: 序号
    
    输出变量：
    - success: 是否成功
    - mcp_prompt: 生成的提示词
    - error: 错误信息（如果失败）
    """
    import sys
    import os
    
    # 添加 Python 路径
    workflow_dir = r'C:\Users\FH\source\repos\Auto Finan\Auto Finan\LLM_Integration'
    if workflow_dir not in sys.path:
        sys.path.insert(0, workflow_dir)
    
    try:
        from workflow_core import process_excel_to_mcp_direct, process_excel_to_stage_prompts
    except ImportError as e:
        return {
            "success": False,
            "error": f"无法导入 workflow_core 模块: {str(e)}",
            "suggestion": "请确保 workflow_core.py 在正确路径"
        }
    
    # 从上游节点获取参数
    # 注意：在 Dify 代码节点中，使用 inputs 字典访问变量
    # 不要使用 {{#workflow.variable_name#}} 语法（那是用于其他节点的）
    excel_path = inputs.get('excel_path', '')
    sheet_name = inputs.get('sheet_name', '')
    serial = inputs.get('serial', '')
    
    # 验证参数
    if not excel_path or not sheet_name or not serial:
        return {
            "success": False,
            "error": "缺少必要参数",
            "required": ["excel_path", "sheet_name", "serial"]
        }
    
    # 检查文件是否存在
    if not os.path.exists(excel_path):
        return {
            "success": False,
            "error": f"Excel 文件不存在: {excel_path}"
        }
    
    try:
        result = process_excel_to_stage_prompts(
            excel_path=excel_path,
            sheet_name=sheet_name,
            serial=serial
        )

        if not result:
            return {
                "success": False,
                "error": "未能生成有效的 MCP 提示词",
                "suggestion": "请检查 Excel 数据和序号是否正确"
            }

        full_prompt = result.get("full_prompt") or process_excel_to_mcp_direct(
            excel_path=excel_path,
            sheet_name=sheet_name,
            serial=serial
        )

        if not full_prompt:
            return {
                "success": False,
                "error": "未能生成完整提示词",
                "suggestion": "请检查 Excel 数据和序号是否正确"
            }

        return {
            "success": True,
            "mcp_prompt": full_prompt,
            "full_prompt": full_prompt,
            "prompt_length": len(full_prompt),
            "stage_prompts": result.get("stage_prompts", {}),
            "business_type": result.get("business_type"),
            "excel_path": excel_path,
            "sheet_name": sheet_name,
            "serial": serial
        }

    except Exception as e:
        return {
            "success": False,
            "error": f"生成提示词失败: {str(e)}",
            "error_type": type(e).__name__
        }


# ============================================================================
# 节点 2：处理 HTTP 响应
# ============================================================================

def node_process_http_response():
    """
    Dify 代码节点：处理 Playwright MCP HTTP 响应
    
    输入变量：
    - http_response: HTTP 请求节点的响应
    
    输出变量：
    - success: 是否成功
    - status: 执行状态
    - message: 消息
    - logs: 日志列表
    """
    import json
    
    # 从 HTTP 请求节点获取响应
    # 注意：在代码节点中，使用 inputs 字典访问变量
    http_response = inputs.get('http_response', {})
    
    # 检查响应格式
    if isinstance(http_response, str):
        try:
            http_response = json.loads(http_response)
        except:
            return {
                "success": False,
                "error": "HTTP 响应格式错误",
                "raw_response": http_response[:500] if len(str(http_response)) > 500 else http_response
            }
    
    # 检查执行状态
    status = http_response.get("status", "unknown")
    message = http_response.get("message", "")
    logs = http_response.get("logs", [])
    execution_id = http_response.get("execution_id", "")
    
    if status == "success":
        return {
            "success": True,
            "status": status,
            "message": message,
            "execution_id": execution_id,
            "logs": logs,
            "logs_count": len(logs),
            "note": "执行成功"
        }
    elif status == "partial":
        return {
            "success": True,
            "status": status,
            "message": message,
            "execution_id": execution_id,
            "logs": logs,
            "warning": "部分步骤执行成功",
            "note": "请检查失败步骤"
        }
    else:
        # 执行失败
        error_details = http_response.get("error_details", {})
        return {
            "success": False,
            "status": status,
            "message": message,
            "error": error_details,
            "logs": logs,
            "note": "执行失败，请检查错误信息"
        }


# ============================================================================
# 节点 3：批量处理（可选）
# ============================================================================

def node_batch_process():
    """
    Dify 代码节点：批量处理 Excel 所有序号
    
    输入变量：
    - excel_path: Excel 文件路径
    - sheet_name: 工作表名称
    
    输出变量：
    - success: 是否成功
    - results: 结果列表
    - count: 数量
    """
    import sys
    import os
    
    # 添加 Python 路径
    workflow_dir = r'C:\Users\FH\source\repos\Auto Finan\Auto Finan\LLM_Integration'
    if workflow_dir not in sys.path:
        sys.path.insert(0, workflow_dir)
    
    try:
        from workflow_core import batch_process_excel_to_mcp_direct
    except ImportError as e:
        return {
            "success": False,
            "error": f"无法导入 workflow_core 模块: {str(e)}"
        }
    
    # 从上游节点获取参数
    # 注意：在代码节点中，使用 inputs 字典访问变量
    excel_path = inputs.get('excel_path', '')
    sheet_name = inputs.get('sheet_name', '')
    
    if not excel_path or not sheet_name:
        return {
            "success": False,
            "error": "缺少必要参数"
        }
    
    if not os.path.exists(excel_path):
        return {
            "success": False,
            "error": f"Excel 文件不存在: {excel_path}"
        }
    
    try:
        # 批量生成所有序号的提示词
        results = batch_process_excel_to_mcp_direct(excel_path, sheet_name)
        
        return {
            "success": True,
            "results": results,
            "count": len(results),
            "message": f"共生成 {len(results)} 个提示词"
        }
        
    except Exception as e:
        return {
            "success": False,
            "error": f"批量处理失败: {str(e)}",
            "error_type": type(e).__name__
        }


# ============================================================================
# 使用说明
# ============================================================================

"""
在 Dify 中使用这些代码：

1. 复制对应的函数代码到 Dify 代码节点
2. 替换 {{#workflow.variable_name#}} 为实际的变量引用
3. 确保 Python 环境可以访问 workflow_core 模块
4. 或者将 workflow_core 的代码直接内联到节点中

示例（在 Dify 代码节点中）：
"""

EXAMPLE_DIFY_CODE = """
import sys
import os

# 添加路径
sys.path.insert(0, r'C:\\Users\\FH\\source\\repos\\Auto Finan\\Auto Finan\\LLM_Integration')

from workflow_core import process_excel_to_mcp_direct

# 获取输入变量（Dify 会自动替换）
excel_path = {{#workflow.excel_path#}}
sheet_name = {{#workflow.sheet_name#}}
serial = {{#workflow.serial#}}

# 生成提示词
mcp_prompt = process_excel_to_mcp_direct(excel_path, sheet_name, serial)

# 返回结果
return {
    "success": True,
    "mcp_prompt": mcp_prompt
}
"""

if __name__ == "__main__":
    print("这是 Dify 工作流代码节点示例")
    print("请复制对应的函数代码到 Dify 代码节点中使用")

