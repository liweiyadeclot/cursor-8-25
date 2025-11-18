#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Playwright 直接执行器

功能：
1. 解析 MCP 提示词
2. 转换为 Playwright 代码
3. 直接执行浏览器操作
4. 返回执行结果

这样就不需要通过 MCP SSE 接口，可以直接执行。
"""

import os
import sys
import re
import json
from typing import Dict, Any, List, Optional
from datetime import datetime

try:
    from playwright.sync_api import sync_playwright, Page, Browser
except ImportError:
    print("❌ 缺少 playwright，请安装: pip install playwright")
    print("   然后运行: playwright install chromium")
    sys.exit(1)


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
        
        if line and "请你调用Playwright MCP" not in line:
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
            return True
        
        # 输入操作
        elif "输入框中输入" in step:
            match = re.search(r'在(.+?)输入框中输入(.+)', step)
            if match:
                label, value = match.groups()
                logs.append(f"正在在 {label} 输入框中输入: {value}")
                # 尝试通过 label 查找输入框
                try:
                    input_selector = f"label:has-text('{label}') + input, input[placeholder*='{label}'], input[name*='{label}']"
                    page.fill(input_selector, value.strip(), timeout=5000)
                    logs.append(f"✅ 已输入: {value}")
                    return True
                except:
                    # 如果失败，尝试其他方式
                    logs.append(f"⚠️  无法找到输入框: {label}")
                    return False
            return False
        
        # 下拉框选择
        elif "下拉框中选择" in step:
            match = re.search(r'在(.+?)下拉框中选择(.+)', step)
            if match:
                label, value = match.groups()
                value = value.strip().strip('"')
                logs.append(f"正在在 {label} 下拉框中选择: {value}")
                try:
                    select_selector = f"label:has-text('{label}') + select, select[name*='{label}']"
                    page.select_option(select_selector, value, timeout=5000)
                    logs.append(f"✅ 已选择: {value}")
                    return True
                except:
                    logs.append(f"⚠️  无法找到下拉框: {label}")
                    return False
            return False
        
        # 点击按钮
        elif "按钮" in step and "点击" in step:
            match = re.search(r'点击(.+?)按钮', step)
            if match:
                button_text = match.group(1)
                logs.append(f"正在点击按钮: {button_text}")
                try:
                    page.click(f"button:has-text('{button_text}'), a:has-text('{button_text}')", timeout=5000)
                    logs.append(f"✅ 已点击: {button_text}")
                    page.wait_for_timeout(1000)  # 等待页面响应
                    return True
                except:
                    logs.append(f"⚠️  无法找到按钮: {button_text}")
                    return False
            return False
        
        # 填写操作
        elif "填写" in step:
            match = re.search(r'向(.+?)输入框填写(.+)', step)
            if match:
                label, value = match.groups()
                logs.append(f"正在向 {label} 输入框填写: {value}")
                try:
                    input_selector = f"label:has-text('{label}') + input, input[placeholder*='{label}']"
                    page.fill(input_selector, value.strip(), timeout=5000)
                    logs.append(f"✅ 已填写: {value}")
                    return True
                except:
                    logs.append(f"⚠️  无法找到输入框: {label}")
                    return False
            return False
        
        # 等待
        elif "等待" in step:
            logs.append(f"等待: {step}")
            page.wait_for_timeout(2000)
            return True
        
        # 其他操作（如保存图片、运行脚本等）
        else:
            logs.append(f"⚠️  未识别的操作: {step}")
            logs.append("💡 提示：某些操作（如保存图片、运行脚本）需要特殊处理")
            return True  # 不阻止后续操作
        
    except Exception as e:
        logs.append(f"❌ 执行失败: {step}")
        logs.append(f"   错误: {str(e)}")
        return False


def execute_playwright_prompt(prompt: str, headless: bool = False, browser_type: str = "chromium") -> Dict[str, Any]:
    """
    执行 Playwright 提示词
    
    Args:
        prompt: MCP 提示词字符串
        headless: 是否无头模式
        browser_type: 浏览器类型 (chromium, firefox, webkit)
    
    Returns:
        执行结果
    """
    execution_id = f"exec_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    logs = []
    steps = parse_mcp_prompt(prompt)
    
    logs.append(f"开始执行，共 {len(steps)} 个步骤")
    
    try:
        with sync_playwright() as p:
            # 启动浏览器
            if browser_type == "chromium":
                browser = p.chromium.launch(headless=headless)
            elif browser_type == "firefox":
                browser = p.firefox.launch(headless=headless)
            elif browser_type == "webkit":
                browser = p.webkit.launch(headless=headless)
            else:
                browser = p.chromium.launch(headless=headless)
            
            logs.append(f"✅ 浏览器已启动: {browser_type}")
            
            # 创建页面
            page = browser.new_page()
            logs.append("✅ 新页面已创建")
            
            # 执行每个步骤
            success_count = 0
            failed_steps = []
            
            for i, step in enumerate(steps, 1):
                logs.append(f"\n--- 步骤 {i}/{len(steps)} ---")
                if execute_step(page, step, logs):
                    success_count += 1
                else:
                    failed_steps.append((i, step))
            
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


if __name__ == "__main__":
    # 测试
    test_prompt = """请你调用Playwright MCP，执行以下命令，一次性执行完
打开https://example.com
在搜索输入框中输入test
点击搜索按钮"""
    
    print("=" * 80)
    print("Playwright 直接执行器测试")
    print("=" * 80)
    
    result = execute_playwright_prompt(test_prompt, headless=False)
    
    print("\n执行结果:")
    print(json.dumps(result, indent=2, ensure_ascii=False))

