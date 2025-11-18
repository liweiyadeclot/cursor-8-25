#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Playwright MCP HTTP 网关

功能：
1. 接收标准化的 HTTP POST 请求（包含 prompt 字段）
2. 将 prompt 转换为 Playwright MCP 可以执行的格式
3. 通过 SSE (Server-Sent Events) 与 Playwright MCP 通信
4. 返回执行结果

使用方法：
    python playwright_mcp_http_gateway.py

或者使用 uvicorn:
    uvicorn playwright_mcp_http_gateway:app --host 0.0.0.0 --port 3030
"""

import os
import json
import asyncio
import subprocess
import re
from typing import Dict, Any, Optional
from datetime import datetime

try:
    from fastapi import FastAPI, HTTPException, Request
    from fastapi.responses import JSONResponse, StreamingResponse
    from pydantic import BaseModel
except ImportError:
    print("❌ 缺少依赖，请安装: pip install fastapi uvicorn")
    raise

# 配置
PLAYWRIGHT_MCP_COMMAND = ["npx", "@playwright/mcp@0.0.46"]
PLAYWRIGHT_MCP_PORT = int(os.environ.get("PLAYWRIGHT_MCP_PORT", "3031"))
GATEWAY_PORT = int(os.environ.get("GATEWAY_PORT", "3030"))
GATEWAY_HOST = os.environ.get("GATEWAY_HOST", "0.0.0.0")

app = FastAPI(title="Playwright MCP HTTP Gateway")


class MCPRequest(BaseModel):
    """MCP 请求模型"""
    prompt: str
    timeout: Optional[int] = 300
    browser: Optional[str] = "chrome"
    headless: Optional[bool] = True


class MCPResponse(BaseModel):
    """MCP 响应模型"""
    status: str
    message: str
    execution_id: Optional[str] = None
    logs: Optional[list] = None
    error_details: Optional[Dict[str, Any]] = None
    timestamp: str


def parse_mcp_prompt(prompt: str) -> list:
    """
    解析 MCP 提示词，提取操作步骤
    
    格式示例：
    1. 请你调用Playwright MCP，执行以下命令，一次性执行完
    2. 打开https://example.com
    3. 在用户名输入框中输入test
    """
    lines = prompt.strip().split('\n')
    steps = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # 移除行首序号（如 "1. "、"2. "）
        line = re.sub(r'^\d+\.\s*', '', line)
        
        if line:
            steps.append(line)
    
    return steps


def extract_playwright_commands(steps: list) -> list:
    """
    从步骤中提取 Playwright 命令
    
    返回格式化的命令列表，可以直接用于 Playwright MCP
    """
    commands = []
    
    for step in steps:
        # 跳过说明性文字
        if "请你调用Playwright MCP" in step or "业务大类" in step:
            continue
        
        # 提取操作类型和参数
        if step.startswith("打开"):
            url = step.replace("打开", "").strip()
            commands.append({"type": "navigate", "url": url})
        
        elif "输入框中输入" in step:
            # 格式：在{控件名}输入框中输入{值}
            match = re.search(r'在(.+?)输入框中输入(.+)', step)
            if match:
                label, value = match.groups()
                commands.append({"type": "fill", "selector": f"label:has-text('{label}')", "value": value.strip()})
        
        elif "下拉框中选择" in step:
            # 格式：在{控件名}下拉框中选择{值}
            match = re.search(r'在(.+?)下拉框中选择(.+)', step)
            if match:
                label, value = match.groups()
                commands.append({"type": "select", "selector": f"label:has-text('{label}')", "value": value.strip().strip('"')})
        
        elif "按钮" in step and "点击" in step:
            # 格式：点击{按钮名}按钮
            match = re.search(r'点击(.+?)按钮', step)
            if match:
                button_text = match.group(1)
                commands.append({"type": "click", "selector": f"button:has-text('{button_text}')"})
        
        elif "选择日期" in step:
            # 格式：选择日期{控件名}为{日期}
            match = re.search(r'选择日期(.+?)为(.+)', step)
            if match:
                label, date = match.groups()
                commands.append({"type": "fill", "selector": f"label:has-text('{label}')", "value": date.strip()})
        
        elif "填写" in step:
            # 格式：向{控件名}输入框填写{值}
            match = re.search(r'向(.+?)输入框填写(.+)', step)
            if match:
                label, value = match.groups()
                commands.append({"type": "fill", "selector": f"label:has-text('{label}')", "value": value.strip()})
        
        elif "等待" in step:
            commands.append({"type": "wait", "timeout": 2000})
        
        else:
            # 其他未识别的命令，作为文本指令传递
            commands.append({"type": "custom", "instruction": step})
    
    return commands


async def execute_playwright_mcp(prompt: str, timeout: int = 300) -> Dict[str, Any]:
    """
    执行 Playwright MCP 命令
    
    注意：Playwright MCP 使用 SSE 通信，这里提供一个简化的 HTTP 网关
    实际执行需要将 prompt 转换为 Playwright MCP 可以理解的格式
    """
    execution_id = f"exec_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    try:
        # 解析提示词
        steps = parse_mcp_prompt(prompt)
        
        # 提取关键信息
        logs = []
        commands_count = 0
        
        for i, step in enumerate(steps, 1):
            step = step.strip()
            if not step:
                continue
            
            # 跳过说明性文字
            if "请你调用Playwright MCP" in step or "业务大类" in step:
                continue
            
            commands_count += 1
            
            # 记录操作
            if "打开" in step:
                url = step.replace("打开", "").strip()
                logs.append(f"步骤 {commands_count}: 打开页面 {url}")
            elif "输入" in step:
                logs.append(f"步骤 {commands_count}: {step}")
            elif "点击" in step:
                logs.append(f"步骤 {commands_count}: {step}")
            elif "选择" in step:
                logs.append(f"步骤 {commands_count}: {step}")
            elif "等待" in step:
                logs.append(f"步骤 {commands_count}: {step}")
            else:
                logs.append(f"步骤 {commands_count}: {step}")
        
        # 注意：这里返回的是解析后的结果
        # 实际执行需要连接到 Playwright MCP 的 SSE 接口
        # 或者通过 Cursor 的 MCP 客户端来执行
        
        return {
            "status": "success",
            "message": "提示词已解析，共识别到 {} 个操作步骤".format(commands_count),
            "execution_id": execution_id,
            "logs": logs,
            "commands_count": commands_count,
            "steps": steps,
            "note": "这是解析结果。实际执行需要通过 Cursor 的 MCP 客户端或 Playwright MCP SSE 接口。"
        }
        
    except Exception as e:
        return {
            "status": "error",
            "message": f"执行失败: {str(e)}",
            "execution_id": execution_id,
            "error_details": {
                "error_type": type(e).__name__,
                "error_message": str(e)
            }
        }


@app.get("/")
async def root():
    """根路径，返回服务信息"""
    return {
        "service": "Playwright MCP HTTP Gateway",
        "version": "1.0.0",
        "status": "running",
        "endpoints": {
            "execute": "/mcp/execute",
            "health": "/health"
        }
    }


@app.get("/health")
async def health():
    """健康检查"""
    return {
        "status": "healthy",
        "timestamp": datetime.now().isoformat()
    }


@app.post("/mcp/execute", response_model=MCPResponse)
async def execute_mcp(request: MCPRequest):
    """
    执行 Playwright MCP 命令
    
    请求体：
    {
        "prompt": "1. 请你调用Playwright MCP...",
        "timeout": 300,
        "browser": "chrome",
        "headless": true
    }
    """
    if not request.prompt or not request.prompt.strip():
        raise HTTPException(status_code=400, detail="prompt 字段不能为空")
    
    try:
        result = await execute_playwright_mcp(
            prompt=request.prompt,
            timeout=request.timeout or 300
        )
        
        result["timestamp"] = datetime.now().isoformat()
        return JSONResponse(content=result)
        
    except Exception as e:
        error_response = {
            "status": "error",
            "message": f"服务器错误: {str(e)}",
            "timestamp": datetime.now().isoformat(),
            "error_details": {
                "error_type": type(e).__name__,
                "error_message": str(e)
            }
        }
        return JSONResponse(status_code=500, content=error_response)


@app.post("/mcp/execute/simple")
async def execute_mcp_simple(prompt: str):
    """
    简化版执行接口，直接接收 prompt 字符串
    
    请求体：纯文本或 JSON {"prompt": "..."}
    """
    if not prompt or not prompt.strip():
        raise HTTPException(status_code=400, detail="prompt 不能为空")
    
    request = MCPRequest(prompt=prompt)
    return await execute_mcp(request)


if __name__ == "__main__":
    import uvicorn
    
    print("="*80)
    print("Playwright MCP HTTP Gateway")
    print("="*80)
    print(f"🌐 服务地址: http://{GATEWAY_HOST}:{GATEWAY_PORT}")
    print(f"📡 执行端点: http://{GATEWAY_HOST}:{GATEWAY_PORT}/mcp/execute")
    print(f"❤️  健康检查: http://{GATEWAY_HOST}:{GATEWAY_PORT}/health")
    print("="*80)
    print()
    
    uvicorn.run(
        app,
        host=GATEWAY_HOST,
        port=GATEWAY_PORT,
        log_level="info"
    )

