#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Playwright MCP HTTP 网关（真正执行版本）

功能：
1. 接收标准化的 HTTP POST 请求（包含 prompt 字段）
2. 解析提示词，提取操作步骤
3. 使用 Playwright 直接执行浏览器操作
4. 返回执行结果

使用方法：
    python playwright_mcp_http_gateway_executor.py

或者使用 uvicorn:
    uvicorn playwright_mcp_http_gateway_executor:app --host 0.0.0.0 --port 3030
"""

import os
import sys
import json
import re
import subprocess
import logging
import traceback
import asyncio
from typing import Dict, Any, Optional, List
from datetime import datetime

try:
    from fastapi import FastAPI, HTTPException, Request
    from fastapi.responses import JSONResponse
    from pydantic import BaseModel
except ImportError:
    print("❌ 缺少依赖，请安装: pip install fastapi uvicorn")
    sys.exit(1)

try:
    from playwright.sync_api import sync_playwright, Page, Browser
except ImportError:
    print("❌ 缺少 playwright，请安装: pip install playwright")
    print("   然后运行: playwright install chromium")
    sys.exit(1)

# 配置
GATEWAY_PORT = int(os.environ.get("GATEWAY_PORT", "3030"))
GATEWAY_HOST = os.environ.get("GATEWAY_HOST", "0.0.0.0")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s"
)
logger = logging.getLogger("playwright_gateway")

app = FastAPI(title="Playwright MCP HTTP Gateway (Executor)")


class MCPRequest(BaseModel):
    """MCP 请求模型"""
    prompt: str
    timeout: Optional[int] = 300
    browser: Optional[str] = "chrome"
    headless: Optional[bool] = False  # 默认显示浏览器，方便调试


class MCPResponse(BaseModel):
    """MCP 响应模型"""
    status: str
    message: str
    execution_id: Optional[str] = None
    logs: Optional[list] = None
    error_details: Optional[Dict[str, Any]] = None
    timestamp: str


def parse_mcp_prompt(prompt: str) -> List[str]:
    """解析 MCP 提示词，提取操作步骤"""
    lines = prompt.strip().split('\n')
    steps = []
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        # 移除行首序号（如 "1. "、"2. "）
        line = re.sub(r'^\d+\.\s*', '', line)
        
        # 跳过说明性文字
        if line and "请你调用Playwright MCP" not in line and "业务大类" not in line:
            steps.append(line)
    
    return steps


def execute_step(page: Page, step: str, logs: List[str]) -> bool:
    """
    执行单个操作步骤
    
    Returns:
        是否成功
    """
    step = step.strip()
    if not step:
        return True
    
    try:
        # 打开页面
        if step.startswith("打开"):
            url = step.replace("打开", "").strip()
            logs.append(f"正在打开页面: {url}")
            page.goto(url, wait_until="networkidle", timeout=60000)
            logs.append(f"✅ 页面已打开: {url}")
            page.wait_for_timeout(1000)  # 等待页面稳定
            return True
        
        # 输入操作
        elif "输入框中输入" in step:
            match = re.search(r'在(.+?)输入框中输入(.+)', step)
            if match:
                label, value = match.groups()
                logs.append(f"正在在 {label} 输入框中输入: {value.strip()}")
                # 尝试多种选择器
                selectors = [
                    f"label:has-text('{label}') + input",
                    f"input[placeholder*='{label}']",
                    f"input[name*='{label}']",
                    f"input[id*='{label}']",
                    f"//label[contains(text(), '{label}')]/following-sibling::input[1]",
                    f"//label[contains(text(), '{label}')]/../input"
                ]
                for selector in selectors:
                    try:
                        if selector.startswith("//"):
                            page.locator(selector).fill(value.strip(), timeout=5000)
                        else:
                            page.fill(selector, value.strip(), timeout=5000)
                        logs.append(f"✅ 已输入: {value.strip()}")
                        return True
                    except:
                        continue
                logs.append(f"⚠️  无法找到输入框: {label}")
                return False
            return False
        
        # 下拉框选择
        elif "下拉框中选择" in step:
            match = re.search(r'在(.+?)下拉框中选择值为(.+)', step)
            if not match:
                match = re.search(r'在(.+?)下拉框中选择(.+)', step)
            if match:
                label, value = match.groups()
                value = value.strip().strip('"').strip("'")
                logs.append(f"正在在 {label} 下拉框中选择: {value}")
                selectors = [
                    f"label:has-text('{label}') + select",
                    f"select[name*='{label}']",
                    f"select[id*='{label}']",
                    f"//label[contains(text(), '{label}')]/following-sibling::select[1]"
                ]
                for selector in selectors:
                    try:
                        if selector.startswith("//"):
                            page.locator(selector).select_option(value, timeout=5000)
                        else:
                            page.select_option(selector, value, timeout=5000)
                        logs.append(f"✅ 已选择: {value}")
                        page.wait_for_timeout(500)  # 等待选择生效
                        return True
                    except:
                        continue
                logs.append(f"⚠️  无法找到下拉框: {label}")
                return False
            return False
        
        # 点击按钮
        elif "按钮" in step and "点击" in step:
            match = re.search(r'点击(.+?)按钮', step)
            if match:
                button_text = match.group(1)
                logs.append(f"正在点击按钮: {button_text}")
                selectors = [
                    f"button:has-text('{button_text}')",
                    f"a:has-text('{button_text}')",
                    f"input[type='button'][value='{button_text}']",
                    f"input[type='submit'][value='{button_text}']",
                    f"//button[contains(text(), '{button_text}')]",
                    f"//a[contains(text(), '{button_text}')]"
                ]
                for selector in selectors:
                    try:
                        if selector.startswith("//"):
                            page.locator(selector).click(timeout=5000)
                        else:
                            page.click(selector, timeout=5000)
                        logs.append(f"✅ 已点击: {button_text}")
                        page.wait_for_timeout(2000)  # 等待页面响应
                        return True
                    except:
                        continue
                logs.append(f"⚠️  无法找到按钮: {button_text}")
                return False
            return False
        
        # 填写操作
        elif "填写" in step:
            match = re.search(r'向(.+?)输入框填写(.+)', step)
            if match:
                label, value = match.groups()
                logs.append(f"正在向 {label} 输入框填写: {value.strip()}")
                selectors = [
                    f"label:has-text('{label}') + input",
                    f"input[placeholder*='{label}']",
                    f"input[name*='{label}']"
                ]
                for selector in selectors:
                    try:
                        page.fill(selector, value.strip(), timeout=5000)
                        logs.append(f"✅ 已填写: {value.strip()}")
                        return True
                    except:
                        continue
                logs.append(f"⚠️  无法找到输入框: {label}")
                return False
            return False
        
        # 选择日期
        elif "选择日期" in step:
            match = re.search(r'选择日期(.+?)为(.+)', step)
            if match:
                label, date = match.groups()
                logs.append(f"正在选择日期 {label}: {date.strip()}")
                selectors = [
                    f"label:has-text('{label}') + input[type='date']",
                    f"input[type='date'][name*='{label}']",
                    f"input[type='text'][name*='{label}']"
                ]
                for selector in selectors:
                    try:
                        page.fill(selector, date.strip(), timeout=5000)
                        logs.append(f"✅ 已选择日期: {date.strip()}")
                        return True
                    except:
                        continue
                logs.append(f"⚠️  无法找到日期选择器: {label}")
                return False
            return False
        
        # 银行卡号尾号
        elif "银行卡号尾号" in step:
            match = re.search(r'银行卡号尾号内容为(.+)', step)
            if match:
                tail = match.groups()[0].strip()
                logs.append(f"正在输入银行卡号尾号: {tail}")
                # 尝试找到银行卡号输入框
                selectors = [
                    "input[name*='card']",
                    "input[name*='bank']",
                    "input[placeholder*='卡号']",
                    "input[placeholder*='尾号']"
                ]
                for selector in selectors:
                    try:
                        page.fill(selector, tail, timeout=5000)
                        logs.append(f"✅ 已输入银行卡号尾号: {tail}")
                        return True
                    except:
                        continue
                logs.append(f"⚠️  无法找到银行卡号输入框")
                return False
        
        # 保存验证码图片
        elif "保存" in step and ("验证码" in step or "图片" in step):
            # 格式：将验证码图片保存至...目录下，命名为...
            match = re.search(r'保存至(.+?)(?:目录下)?，命名为(.+)', step)
            if match:
                save_dir, filename = match.groups()
                save_dir = save_dir.strip()
                filename = filename.strip()
                logs.append(f"正在保存验证码图片到: {save_dir}/{filename}")
                try:
                    # 查找验证码图片（尝试多种选择器）
                    img_selectors = [
                        "img[src*='captcha']",
                        "img[src*='code']",
                        "img[alt*='验证码']",
                        "img[id*='captcha']",
                        "img[id*='code']",
                        "//img[contains(@src, 'captcha')]",
                        "//img[contains(@src, 'code')]"
                    ]
                    
                    img_found = False
                    for selector in img_selectors:
                        try:
                            if selector.startswith("//"):
                                img_element = page.locator(selector).first
                            else:
                                img_element = page.locator(selector).first
                            
                            if img_element.count() > 0:
                                save_path = os.path.join(save_dir, filename)
                                os.makedirs(save_dir, exist_ok=True)
                                img_element.screenshot(path=save_path)
                                logs.append(f"✅ 验证码图片已保存: {save_path}")
                                img_found = True
                                return True
                        except:
                            continue
                    
                    if not img_found:
                        logs.append(f"⚠️  未找到验证码图片")
                        return False
                except Exception as e:
                    logs.append(f"⚠️  保存验证码图片失败: {str(e)}")
                    return False
        
        # 运行脚本
        elif "运行" in step and ".py" in step:
            # 格式：运行...目录下的OCR.py
            match = re.search(r'运行(.+?\.py)', step)
            if match:
                script_path = match.group(1).strip()
                # 处理路径中的反斜杠
                script_path = script_path.replace('\\', os.sep).replace('/', os.sep)
                logs.append(f"正在运行脚本: {script_path}")
                try:
                    # 检查文件是否存在
                    if not os.path.exists(script_path):
                        logs.append(f"⚠️  脚本文件不存在: {script_path}")
                        return False
                    
                    result = subprocess.run(
                        ["python", script_path],
                        capture_output=True,
                        text=True,
                        timeout=30,
                        cwd=os.path.dirname(script_path) if os.path.dirname(script_path) else None
                    )
                    if result.returncode == 0:
                        logs.append(f"✅ 脚本执行成功")
                        if result.stdout:
                            logs.append(f"输出: {result.stdout[:200]}")
                        return True
                    else:
                        logs.append(f"⚠️  脚本执行失败: {result.stderr[:200] if result.stderr else '无错误信息'}")
                        return False
                except subprocess.TimeoutExpired:
                    logs.append(f"⚠️  脚本执行超时")
                    return False
                except Exception as e:
                    logs.append(f"⚠️  运行脚本失败: {str(e)}")
                    return False
        
        # 等待
        elif "等待" in step:
            if "页面响应" in step or "页面跳转" in step:
                logs.append(f"等待页面响应...")
                page.wait_for_timeout(3000)
            else:
                logs.append(f"等待: {step}")
                page.wait_for_timeout(2000)
            return True
        
        # 调用脚本（带参数）
        elif "调用" in step and ".py" in step:
            # 格式：调用test_mouse_keyboard.py，执行一个python自动点击的脚本，脚本的第一个参数为保存路径，第二个参数为保存文件名
            logs.append(f"⚠️  调用脚本操作: {step}")
            logs.append(f"💡 提示：此操作需要特殊处理，当前版本暂不支持")
            # TODO: 实现脚本调用逻辑
            return True  # 不阻止后续操作
        
        # 重命名文件
        elif "重命名" in step:
            logs.append(f"⚠️  文件重命名操作: {step}")
            logs.append(f"💡 提示：此操作需要特殊处理，当前版本暂不支持")
            # TODO: 实现文件重命名逻辑
            return True  # 不阻止后续操作
        
        # 其他未识别的操作
        else:
            logs.append(f"⚠️  未识别的操作: {step}")
            logs.append(f"💡 提示：某些操作可能需要特殊处理")
            return True  # 不阻止后续操作
        
    except Exception as e:
        logs.append(f"❌ 执行失败: {step}")
        logs.append(f"   错误: {str(e)}")
        return False


def execute_playwright_prompt(prompt: str, headless: bool = False, browser_type: str = "chromium", timeout: int = 300) -> Dict[str, Any]:
    """
    执行 Playwright 提示词
    
    Args:
        prompt: MCP 提示词字符串
        headless: 是否无头模式
        browser_type: 浏览器类型 (chromium, firefox, webkit)
        timeout: 总超时时间（秒）
    
    Returns:
        执行结果
    """
    execution_id = f"exec_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    logs = []
    steps = parse_mcp_prompt(prompt)
    
    logs.append(f"🚀 开始执行，共 {len(steps)} 个步骤")
    
    try:
        with sync_playwright() as p:
            # 启动浏览器
            browser_options = {
                "headless": headless,
                "slow_mo": 100  # 减慢操作速度，便于观察
            }
            
            if browser_type == "chromium" or browser_type == "chrome":
                browser = p.chromium.launch(**browser_options)
            elif browser_type == "firefox":
                browser = p.firefox.launch(**browser_options)
            elif browser_type == "webkit":
                browser = p.webkit.launch(**browser_options)
            else:
                browser = p.chromium.launch(**browser_options)
            
            logs.append(f"✅ 浏览器已启动: {browser_type} (headless={headless})")
            
            # 创建页面
            page = browser.new_page()
            logs.append("✅ 新页面已创建")
            
            # 设置超时
            page.set_default_timeout(timeout * 1000)
            
            # 执行每个步骤
            success_count = 0
            failed_steps = []
            
            for i, step in enumerate(steps, 1):
                logs.append(f"\n--- 步骤 {i}/{len(steps)}: {step[:50]}... ---")
                if execute_step(page, step, logs):
                    success_count += 1
                else:
                    failed_steps.append((i, step))
                    # 如果关键步骤失败，可以选择继续或停止
                    # 这里选择继续执行
            
            # 等待一段时间，让用户看到最终结果
            if not headless:
                logs.append("\n⏳ 等待 3 秒后关闭浏览器...")
                page.wait_for_timeout(3000)
            
            # 关闭浏览器
            browser.close()
            logs.append("\n✅ 浏览器已关闭")
            
            # 返回结果
            if len(failed_steps) == 0:
                return {
                    "status": "success",
                    "message": f"所有步骤执行成功（{success_count}/{len(steps)}）",
                    "execution_id": execution_id,
                    "logs": logs,
                    "total_steps": len(steps),
                    "success_steps": success_count,
                    "failed_steps": []
                }
            else:
                return {
                    "status": "partial",
                    "message": f"部分步骤执行成功（{success_count}/{len(steps)}）",
                    "execution_id": execution_id,
                    "logs": logs,
                    "total_steps": len(steps),
                    "success_steps": success_count,
                    "failed_steps": failed_steps
                }
    
    except Exception as e:
        return {
            "status": "error",
            "message": f"执行失败: {str(e)}",
            "execution_id": execution_id,
            "logs": logs,
            "error_details": {
                "error_type": type(e).__name__,
                "error_message": str(e)
            }
        }


@app.get("/")
async def root():
    """根路径，返回服务信息"""
    return {
        "service": "Playwright MCP HTTP Gateway (Executor)",
        "version": "2.0.0",
        "status": "running",
        "capabilities": {
            "real_execution": True,
            "browser_automation": True
        },
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
        "timestamp": datetime.now().isoformat(),
        "capabilities": ["real_execution", "browser_automation"]
    }


@app.post("/mcp/execute")
async def execute_mcp(request: MCPRequest):
    """
    执行 Playwright MCP 命令（真正执行版本）
    
    请求体：
    {
        "prompt": "1. 请你调用Playwright MCP...",
        "timeout": 300,
        "browser": "chrome",
        "headless": false
    }
    """
    if not request.prompt or not request.prompt.strip():
        raise HTTPException(status_code=400, detail="prompt 字段不能为空")
    
    try:
        logger.info("收到执行请求 headless=%s browser=%s timeout=%s", request.headless, request.browser, request.timeout)
        # 转换浏览器类型
        browser_type = request.browser or "chrome"
        if browser_type == "chrome":
            browser_type = "chromium"
        
        # 在后台线程中执行同步 Playwright 逻辑，避免阻塞 asyncio 事件循环
        loop = asyncio.get_running_loop()
        result = await loop.run_in_executor(
            None,
            lambda: execute_playwright_prompt(
                prompt=request.prompt,
                headless=request.headless if request.headless is not None else False,
                browser_type=browser_type,
                timeout=request.timeout or 300
            )
        )
        
        result["timestamp"] = datetime.now().isoformat()
        return JSONResponse(content=result)
        
    except Exception as e:
        logger.error("执行失败: %s", e)
        logger.error(traceback.format_exc())
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


if __name__ == "__main__":
    import uvicorn
    
    print("="*80)
    print("Playwright MCP HTTP Gateway (真正执行版本)")
    print("="*80)
    print(f"🌐 服务地址: http://{GATEWAY_HOST}:{GATEWAY_PORT}")
    print(f"📡 执行端点: http://{GATEWAY_HOST}:{GATEWAY_PORT}/mcp/execute")
    print(f"❤️  健康检查: http://{GATEWAY_HOST}:{GATEWAY_PORT}/health")
    print()
    print("✨ 功能：真正执行浏览器操作，而不仅仅是解析提示词")
    print("="*80)
    print()
    
    uvicorn.run(
        app,
        host=GATEWAY_HOST,
        port=GATEWAY_PORT,
        log_level="info"
    )

