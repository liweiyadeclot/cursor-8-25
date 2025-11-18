#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dify 本地服务

功能：
1. 在本地运行 HTTP 服务
2. 接收 Dify 的 HTTP 请求
3. 调用本地的 workflow_core.py 处理
4. 返回结果给 Dify

这样 Dify 就不需要内联代码，可以直接调用本地的功能。
"""

import os
import sys
from typing import Dict, Any, Optional

# 添加当前目录到路径
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

try:
    from fastapi import FastAPI, HTTPException
    from fastapi.responses import JSONResponse
    from pydantic import BaseModel
    import uvicorn
except ImportError:
    print("❌ 缺少依赖，请安装: pip install fastapi uvicorn")
    sys.exit(1)

try:
    from workflow_core import (
        process_excel_to_mcp_direct,
        batch_process_excel_to_mcp_direct,
        process_excel_to_stage_prompts,
    )
except ImportError as e:
    print(f"❌ 无法导入 workflow_core: {e}")
    sys.exit(1)

# 修复 Windows 控制台编码
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except:
        pass

app = FastAPI(title="Dify Local Service - Excel to MCP Prompt")


class ExcelToPromptRequest(BaseModel):
    """Excel 转提示词请求"""
    excel_path: str
    sheet_name: str
    serial: str
    
    class Config:
        # 允许字段为空（在验证中处理）
        pass


class BatchProcessRequest(BaseModel):
    """批量处理请求"""
    excel_path: str
    sheet_name: str


@app.get("/")
async def root():
    """根路径"""
    return {
        "service": "Dify Local Service",
        "version": "1.0.0",
        "status": "running",
        "endpoints": {
            "excel_to_prompt": "/api/excel-to-prompt",
            "batch_process": "/api/batch-process",
            "health": "/health"
        }
    }


@app.get("/health")
async def health():
    """健康检查"""
    return {
        "status": "healthy",
        "service": "Dify Local Service"
    }


@app.post("/api/excel-to-prompt")
async def excel_to_prompt(request: ExcelToPromptRequest):
    """
    Excel 转 MCP 提示词
    
    请求体：
    {
        "excel_path": "C:\\path\\to\\file.xlsx",
        "sheet_name": "3-报销",
        "serial": "1"
    }
    """
    try:
        # 验证必需字段
        if not request.excel_path or not request.sheet_name or not request.serial:
            return JSONResponse(
                status_code=400,
                content={
                    "success": False,
                    "error": "缺少必需字段",
                    "required": ["excel_path", "sheet_name", "serial"],
                    "received": {
                        "excel_path": request.excel_path,
                        "sheet_name": request.sheet_name,
                        "serial": request.serial
                    }
                }
            )
        # 验证文件是否存在
        if not os.path.exists(request.excel_path):
            raise HTTPException(
                status_code=404,
                detail=f"Excel 文件不存在: {request.excel_path}"
            )
        
        stage_error = None
        result = {}
        try:
            result = process_excel_to_stage_prompts(
                excel_path=request.excel_path,
                sheet_name=request.sheet_name,
                serial=request.serial
            ) or {}
        except Exception as e:
            stage_error = str(e)

        full_prompt = (result.get("full_prompt", "") if isinstance(result, dict) else "") or ""
        stage_prompts = (result.get("stage_prompts", {}) if isinstance(result, dict) else {}) or {}

        if not full_prompt:
            full_prompt = process_excel_to_mcp_direct(
                excel_path=request.excel_path,
                sheet_name=request.sheet_name,
                serial=request.serial
            )

        if not full_prompt or not full_prompt.strip():
            return JSONResponse(
                status_code=400,
                content={
                    "success": False,
                    "error": f"序号 {request.serial} 未生成有效的完整提示词",
                    "suggestion": "请检查 Excel 数据和序号是否正确"
                }
            )

        return {
            "success": True,
            "business_type": result.get("business_type"),
            "full_prompt": full_prompt,
            "prompt_length": len(full_prompt),
            "stage_prompts": stage_prompts,
            "stage_error": stage_error,
            "excel_path": request.excel_path,
            "sheet_name": request.sheet_name,
            "serial": request.serial
        }
        
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "success": False,
                "error": f"处理失败: {str(e)}",
                "error_type": type(e).__name__
            }
        )


@app.post("/api/batch-process")
async def batch_process(request: BatchProcessRequest):
    """
    批量处理 Excel 所有序号
    
    请求体：
    {
        "excel_path": "C:\\path\\to\\file.xlsx",
        "sheet_name": "3-报销"
    }
    """
    try:
        # 验证文件是否存在
        if not os.path.exists(request.excel_path):
            raise HTTPException(
                status_code=404,
                detail=f"Excel 文件不存在: {request.excel_path}"
            )
        
        # 调用本地函数
        results = batch_process_excel_to_mcp_direct(
            excel_path=request.excel_path,
            sheet_name=request.sheet_name
        )
        
        return {
            "success": True,
            "results": results,
            "count": len(results),
            "message": f"共处理 {len(results)} 个序号"
        }
        
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "success": False,
                "error": f"批量处理失败: {str(e)}",
                "error_type": type(e).__name__
            }
        )


if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Dify 本地服务")
    parser.add_argument("--host", default="0.0.0.0", help="绑定主机 (默认: 0.0.0.0)")
    parser.add_argument("--port", type=int, default=8001, help="端口 (默认: 8001)")
    
    args = parser.parse_args()
    
    print("=" * 80)
    print("Dify 本地服务")
    print("=" * 80)
    print(f"🌐 服务地址: http://{args.host}:{args.port}")
    print(f"📡 Excel 转提示词: http://{args.host}:{args.port}/api/excel-to-prompt")
    print(f"📡 批量处理: http://{args.host}:{args.port}/api/batch-process")
    print(f"❤️  健康检查: http://{args.host}:{args.port}/health")
    print("=" * 80)
    print()
    print("💡 提示：")
    print("   1. 确保防火墙允许端口访问")
    print("   2. 如果 Dify 在远程服务器，使用服务器 IP 地址")
    print("   3. 在 Dify 中使用 HTTP 请求节点调用此服务")
    print()
    
    uvicorn.run(
        app,
        host=args.host,
        port=args.port,
        log_level="info"
    )

