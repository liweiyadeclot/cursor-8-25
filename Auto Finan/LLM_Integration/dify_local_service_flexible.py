#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dify 本地服务（灵活版本）

支持更灵活的请求格式，避免 422 错误
"""

import os
import sys
from typing import Dict, Any, Optional

# 添加当前目录到路径
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

try:
    from fastapi import FastAPI, HTTPException, Request
    from fastapi.responses import JSONResponse
    from pydantic import BaseModel, Field
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
    """Excel 转提示词请求（灵活版本）"""
    excel_path: str = Field(..., description="Excel 文件路径")
    sheet_name: str = Field(..., description="工作表名称")
    serial: str = Field(..., description="序号")


@app.post("/api/excel-to-prompt")
async def excel_to_prompt(request: Request):
    """
    Excel 转 MCP 提示词（灵活版本，支持多种请求格式）
    支持从请求体（JSON）或查询参数（URL）读取数据
    """
    try:
        # 优先从查询参数读取（Dify 可能使用查询参数）
        query_params = dict(request.query_params)
        excel_path = query_params.get("excel_path") or query_params.get("excelPath")
        sheet_name = query_params.get("sheet_name") or query_params.get("sheetName")
        serial = query_params.get("serial")
        
        # 如果查询参数为空，尝试从请求体读取
        if not excel_path or not sheet_name or not serial:
            try:
                body = await request.json()
                excel_path = excel_path or body.get("excel_path") or body.get("excelPath")
                sheet_name = sheet_name or body.get("sheet_name") or body.get("sheetName")
                serial = serial or body.get("serial")
            except:
                # 如果请求体也不是 JSON，使用查询参数（可能为空）
                pass
        
        # URL 解码（处理 URL 编码的参数）
        import urllib.parse
        if excel_path:
            excel_path = urllib.parse.unquote(excel_path)
        if sheet_name:
            sheet_name = urllib.parse.unquote(sheet_name)
        if serial:
            serial = urllib.parse.unquote(str(serial))
        
        # 规范化路径（处理转义字符和 Unicode 编码）
        original_path = excel_path  # 保存原始路径用于调试
        if excel_path:
            # 处理 JSON 转义的 Unicode 字符（如 \u8d22）
            # 方法：使用 json.loads 来正确解码 JSON 字符串中的 Unicode 转义
            try:
                import json
                # 将路径作为 JSON 字符串解析（处理 \uXXXX 转义）
                excel_path = json.loads(f'"{excel_path}"')
            except Exception as e:
                # 如果失败，尝试手动处理 Unicode 转义
                try:
                    import re
                    def decode_unicode(match):
                        return chr(int(match.group(1), 16))
                    excel_path = re.sub(r'\\u([0-9a-fA-F]{4})', decode_unicode, excel_path)
                except:
                    pass
            
            # 彻底处理反斜杠（可能需要多次替换）
            # 因为 JSON 字符串中的 \\\\ 会被解码为 \\，需要继续处理
            iteration = 0
            while '\\\\' in excel_path and iteration < 10:  # 最多10次，防止无限循环
                excel_path = excel_path.replace('\\\\', '\\')
                iteration += 1
            # 处理单反斜杠转义
            excel_path = excel_path.replace('\\/', '/')
            
            # 规范化路径分隔符
            excel_path = os.path.normpath(excel_path)
        
        # 验证必需字段
        if not excel_path:
            return JSONResponse(
                status_code=400,
                content={
                    "success": False,
                    "error": "缺少必需字段: excel_path",
                    "received": body
                }
            )
        
        if not sheet_name:
            return JSONResponse(
                status_code=400,
                content={
                    "success": False,
                    "error": "缺少必需字段: sheet_name",
                    "received": body
                }
            )
        
        if not serial:
            return JSONResponse(
                status_code=400,
                content={
                    "success": False,
                    "error": "缺少必需字段: serial",
                    "received": body
                }
            )
        
        # 验证文件是否存在
        if not os.path.exists(excel_path):
            # 尝试其他可能的路径格式
            alternative_paths = []
            
            # 尝试不同的路径格式
            base_path = excel_path
            # 1. 替换所有反斜杠为正斜杠
            alternative_paths.append(base_path.replace('\\', '/'))
            # 2. 替换所有正斜杠为反斜杠
            alternative_paths.append(base_path.replace('/', '\\'))
            # 3. 再次处理双反斜杠
            temp = base_path
            while '\\\\' in temp:
                temp = temp.replace('\\\\', '\\')
            alternative_paths.append(temp)
            # 4. 使用 os.path.normpath 处理
            alternative_paths.append(os.path.normpath(base_path))
            
            # 去重
            alternative_paths = list(dict.fromkeys(alternative_paths))
            
            found_path = None
            for alt_path in alternative_paths:
                if os.path.exists(alt_path):
                    found_path = alt_path
                    excel_path = alt_path
                    break
            
            if not found_path:
                return JSONResponse(
                    status_code=404,
                    content={
                        "success": False,
                        "error": f"Excel 文件不存在: {excel_path}",
                        "received_path": excel_path,
                        "debug": {
                            "path_exists": False,
                            "path_type": type(excel_path).__name__,
                            "path_length": len(excel_path),
                            "path_repr": repr(excel_path),
                            "tried_alternatives": alternative_paths,
                            "original_request": str(request.url) if hasattr(request, 'url') else None
                        },
                        "suggestion": "请检查文件路径是否正确，确保文件在本地服务可访问的位置。提示：路径中的反斜杠可能需要特殊处理。"
                    }
                )
        
        import traceback
        stage_error = None
        result: Dict[str, Any] = {}
        try:
            result = process_excel_to_stage_prompts(
                excel_path=excel_path,
                sheet_name=sheet_name,
                serial=str(serial)
            ) or {}
        except Exception as e:
            stage_error = str(e)
            if os.environ.get("DEBUG"):
                stage_error = f"{stage_error}\n{traceback.format_exc()}"

        full_prompt = (result.get("full_prompt", "") if isinstance(result, dict) else "") or ""
        stage_prompts = (result.get("stage_prompts", {}) if isinstance(result, dict) else {}) or {}

        if not full_prompt:
            try:
                full_prompt = process_excel_to_mcp_direct(
                    excel_path=excel_path,
                    sheet_name=sheet_name,
                    serial=str(serial)
                )
            except Exception as e:
                return JSONResponse(
                    status_code=500,
                    content={
                        "success": False,
                        "error": f"生成提示词失败: {str(e)}",
                        "error_type": type(e).__name__,
                        "traceback": traceback.format_exc() if os.environ.get("DEBUG") else None,
                        "received": {
                            "excel_path": excel_path,
                            "sheet_name": sheet_name,
                            "serial": str(serial)
                        }
                    }
                )

        if not full_prompt or not full_prompt.strip():
            return JSONResponse(
                status_code=400,
                content={
                    "success": False,
                    "error": f"序号 {serial} 未生成有效的完整提示词",
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
            "excel_path": excel_path,
            "sheet_name": sheet_name,
            "serial": str(serial)
        }
        
    except Exception as e:
        import traceback
        return JSONResponse(
            status_code=500,
            content={
                "success": False,
                "error": f"处理失败: {str(e)}",
                "error_type": type(e).__name__,
                "traceback": traceback.format_exc() if os.environ.get("DEBUG") else None
            }
        )


@app.get("/")
async def root():
    """根路径"""
    return {
        "service": "Dify Local Service (Flexible)",
        "version": "1.0.0",
        "status": "running",
        "endpoints": {
            "excel_to_prompt": "/api/excel-to-prompt",
            "health": "/health"
        },
        "note": "此版本支持更灵活的请求格式，避免 422 错误"
    }


@app.get("/health")
async def health():
    """健康检查"""
    return {
        "status": "healthy",
        "service": "Dify Local Service (Flexible)"
    }


if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Dify 本地服务（灵活版本）")
    parser.add_argument("--host", default="0.0.0.0", help="绑定主机 (默认: 0.0.0.0)")
    parser.add_argument("--port", type=int, default=8001, help="端口 (默认: 8001)")
    parser.add_argument("--debug", action="store_true", help="启用调试模式")
    
    args = parser.parse_args()
    
    if args.debug:
        os.environ["DEBUG"] = "1"
    
    print("=" * 80)
    print("Dify 本地服务（灵活版本）")
    print("=" * 80)
    print(f"🌐 服务地址: http://{args.host}:{args.port}")
    print(f"📡 Excel 转提示词: http://{args.host}:{args.port}/api/excel-to-prompt")
    print(f"❤️  健康检查: http://{args.host}:{args.port}/health")
    print("=" * 80)
    print()
    print("💡 此版本支持更灵活的请求格式，避免 422 验证错误")
    print()
    
    uvicorn.run(
        app,
        host=args.host,
        port=args.port,
        log_level="info"
    )

