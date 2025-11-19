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
import time
import threading
import uuid
import atexit
from typing import Dict, Any, Optional, List, Tuple
from datetime import datetime
from dataclasses import dataclass
from concurrent.futures import ThreadPoolExecutor

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


class LiveLog(list):
    """在追加日志时同步打印到控制台的列表"""

    def append(self, item):  # type: ignore[override]
        text = str(item)
        print(f"[MCP] {text}")
        super().append(text)

    def extend(self, iterable):  # type: ignore[override]
        for item in iterable:
            self.append(item)


PLAYWRIGHT_EXECUTOR = ThreadPoolExecutor(max_workers=1)
_playwright = None
_playwright_lock = threading.Lock()


def get_playwright():
    global _playwright
    with _playwright_lock:
        if _playwright is None:
            _playwright = sync_playwright().start()
    return _playwright


def _shutdown_resources():
    try:
        PLAYWRIGHT_EXECUTOR.shutdown(wait=False, cancel_futures=True)
    except Exception:
        pass
    global _playwright
    if _playwright is not None:
        try:
            _playwright.stop()
        except Exception:
            pass
        _playwright = None


atexit.register(_shutdown_resources)


INPUT_SELECTOR_HINTS = {
    "用户名": ["#loginid", "#zhLoginName", "input[name*='login']", "input[name*='user']", "input[id*='login']"],
    "密码": ["#password", "#zhLoginPwd", "input[type='password']", "input[name*='pwd']"],
    "验证码": ["#checkcode", "input[name*='code']"],
    "报销项目号": [
        "input[id*='uni_prj_code']",
        "input[name*='uni_prj_code']",
        "input[id*='prj_code']",
        "input[name*='prj_code']",
        "input[id*='projectCode']",
        "input[name*='projectCode']",
        "input[id*='project']",
        "input[name*='projectNo']",
        "input[id*='projectNo']"
    ],
    "附件张数": [
        "input[id*='addition']",
        "input[name*='addition']",
        "input[id*='attachment']",
        "input[name*='attachment']",
        "input[id*='fileCount']",
        "input[name*='fileCount']"
    ],
    "出差人": ["input[name*='trav']", "input[name*='person']", "input[id*='trav']"],
    "姓名": ["input[name*='name']", "input[id*='name']"],
    "工作单位": ["input[name*='workUnit']", "input[id*='workUnit']"],
    "职称": ["input[name*='title']", "input[id*='title']"],
    "学工号": ["input[name*='student']", "input[name*='xgh']", "input[id*='student']"],
    "金额": ["input[name*='amount']", "input[id*='amount']"],
    "其他交通费": ["input[name*='transport']", "input[id*='transport']", "input[name*='other']"],
}

DROPDOWN_SELECTOR_HINTS = {
    "人员类型": ["select[name*='personType']", "select[id*='personType']"],
    "省份": ["select[name*='province']", "select[id*='province']"],
    "是否安排伙食": ["select[name*='meal']", "select[id*='meal']"],
    "是否安排交通": ["select[name*='transportArranged']", "select[id*='transportArranged']"],
    "支付方式": [
        "select[id*='pay_type']",
        "select[name*='pay_type']",
        "select[id*='payType']",
        "select[name*='payType']"
    ],
}

BUTTON_SELECTOR_HINTS = {
    "登录": ["#zhLogin", "#loginBtn", "button#zhLogin", "input#zhLogin", "button[name='login']", "input[type='submit'][value*='登录']"],
    "网上预约报账": [
        "div.syslink:has(img[src*='wsyy.jpg'])",
        "div[onclick*='WF_YB6']",
        "div.syslink[title*='点击进入']:has(img[src*='wsyy'])",
        "div:has(img[src*='wsyy.jpg'])",
        "div[onclick*='navToPrj']:has(img[src*='wsyy'])",
        "button:has-text('网上预约报账')",
        "a:has-text('网上预约报账')"
    ],
    "申请报销单": ["button:has-text('申请报销单')", "a:has-text('申请报销单')"],
    "已阅读并同意": ["button:has-text('已阅读并同意')", "input[value*='已阅读']"],
    "下一步": ["button:has-text('下一步')", "input[type='button'][value*='下一步']"],
    "提交": ["button:has-text('提交')", "input[type='submit'][value*='提交']"],
    "预约": ["button:has-text('预约')", "input[value*='预约']"],
    "打印确认单": ["button:has-text('打印确认单')", "input[value*='打印确认单']"],
    "返回": ["button:has-text('返回')", "a:has-text('返回')"],
}

IMAGE_SELECTOR_HINTS = {
    "验证码": ["#checkcodeImg", "img#checkcodeImg", "img[id='checkcodeImg']", "img[src*='CheckCode']", "img[src*='checkcode']"]
}

MAX_STEP_RETRIES = 4
RETRY_INTERVAL_MS = 5000


def _deduplicate_selectors(selectors: List[str]) -> List[str]:
    seen = set()
    ordered = []
    for sel in selectors:
        if sel and sel not in seen:
            ordered.append(sel)
            seen.add(sel)
    return ordered


def _build_selector_candidates(label: str, hint_map: Dict[str, List[str]], defaults: List[str]) -> List[str]:
    selectors: List[str] = []
    for key, values in hint_map.items():
        if key in label:
            selectors.extend(values)
    selectors.extend(hint_map.get(label, []))
    selectors.extend(defaults)
    return _deduplicate_selectors(selectors)


def _extract_id_from_text(text: str) -> Optional[str]:
    match = re.search(r"id[为=]\s*([A-Za-z0-9_\-]+)", text)
    if match:
        return match.group(1).strip()
    return None


def _tokenize_label(label: str) -> List[str]:
    tokens = re.split(r"[^\w]+", label)
    tokens = [t for t in tokens if t]
    lowered = [label.lower()]
    lowered.extend([t.lower() for t in tokens])
    return lowered


def capture_dom_snapshot(page: Page) -> List[Dict[str, Any]]:
    script = """
    (() => {
        const data = [];
        let counter = 1;
        function assignRef(el) {
            const ref = `mcp_${counter++}`;
            el.setAttribute('data-mcp-ref', ref);
            return ref;
        }
        const walker = document.createTreeWalker(document.body || document.documentElement, NodeFilter.SHOW_ELEMENT);
        while (walker.nextNode()) {
            const el = walker.currentNode;
            const ref = assignRef(el);
            data.push({
                ref,
                tag: el.tagName || '',
                id: el.id || '',
                name: el.getAttribute('name') || '',
                type: el.getAttribute('type') || '',
                placeholder: el.getAttribute('placeholder') || '',
                role: el.getAttribute('role') || '',
                text: (el.innerText || '').trim().substring(0, 200)
            });
        }
        return data;
    })();
    """
    try:
        return page.evaluate(script)
    except Exception:
        return []


def _score_node(node: Dict[str, Any], keywords: List[str], label_text: str = "") -> int:
    """
    计算节点与关键词的匹配分数
    
    Args:
        node: 节点数据
        keywords: 关键词列表
        label_text: 标签文本（如"报销项目号"），用于匹配相邻元素的文本
    """
    score = 0
    label_lower = label_text.lower() if label_text else ""
    
    # 获取节点属性
    node_id = (node.get("id") or "").lower()
    node_name = (node.get("name") or "").lower()
    node_placeholder = (node.get("placeholder") or "").lower()
    node_type = (node.get("type") or "").lower()
    node_tag = (node.get("tag") or "").lower()
    
    # 上下文信息
    prev_sibling = (node.get("prevSiblingText") or "").lower()
    row_label = (node.get("rowLabelText") or "").lower()
    parent_text = (node.get("parentText") or "").lower()
    
    # 精确匹配检查：确保不会匹配到错误的字段
    # 例如：查找"密码"时，不应该匹配到"用户名"
    if label_lower:
        # 定义互斥字段（如果查找A，不应该匹配到B）
        exclusive_pairs = [
            (["用户名", "user", "login"], ["密码", "password", "pwd"]),
            (["密码", "password", "pwd"], ["用户名", "user", "login"]),
        ]
        
        for pair_a, pair_b in exclusive_pairs:
            if any(a in label_lower for a in pair_a):
                # 如果查找的是A，检查节点是否明确包含B的关键词
                # 检查 id、name、type、相邻元素、行标签
                node_combined = f"{node_id} {node_name} {node_type} {prev_sibling} {row_label}".lower()
                if any(b in node_combined for b in pair_b):
                    # 但如果节点也包含A的关键词，可能是正确的（例如：login_password 可能是登录密码）
                    if not any(a in node_combined for a in pair_a):
                        return -1000  # 严重不匹配，返回负分
    
    # 1. 精确匹配：标签文本完全匹配在 id 或 name 中（最高优先级）
    if label_lower:
        if label_lower in node_id or node_id in label_lower:
            score += 100
        if label_lower in node_name or node_name in label_lower:
            score += 90
    
    # 2. 关键词匹配：在 id 或 name 中
    for kw in keywords:
        kw_lower = kw.lower()
        if kw_lower in node_id:
            score += 30  # id 中的关键词权重很高
        if kw_lower in node_name:
            score += 25  # name 中的关键词权重高
    
    # 3. 类型匹配：对于特定字段，检查 input type
    if label_lower:
        if "密码" in label_text or "password" in label_lower:
            if node_type == "password":
                score += 50  # 密码输入框必须是 type="password"
            elif node_type == "text":
                score -= 20  # 如果是 text 类型，降低分数
        elif "用户名" in label_text or "user" in label_lower or "login" in label_lower:
            if node_type == "text":
                score += 20  # 用户名通常是 text 类型
            elif node_type == "password":
                score -= 30  # 密码类型不应该是用户名
    
    # 4. 上下文匹配：标签文本在相邻元素或行标签中（高优先级）
    if label_lower and node_tag in ["input", "textarea", "select"]:
        if label_lower in prev_sibling:
            score += 40  # 前一个兄弟元素包含标签文本
        if label_lower in row_label:
            score += 45  # 同一行的标签单元格包含标签文本
        if label_lower in parent_text:
            score += 20  # 父元素包含标签文本
        
        # 特殊处理：radio button 的匹配
        if node_type == "radio" and label_lower:
            # 如果父元素或相邻元素包含标签文本，给予高分
            if label_lower in parent_text:
                score += 50  # 父元素（如 li）包含标签文本，给予高分
            if label_lower in prev_sibling:
                score += 45  # 前一个兄弟元素（如 span）包含标签文本
            # 如果 nextSiblingText 包含标签文本（radio button 后面通常有 span）
            next_sibling = (node.get("nextSiblingText") or "").lower()
            if label_lower in next_sibling:
                score += 45  # 后一个兄弟元素包含标签文本
    
    # 5. placeholder 匹配
    if label_lower and label_lower in node_placeholder:
        score += 15
    
    # 6. 其他属性中的关键词匹配（较低权重）
    combined = " ".join([
        node.get("text") or "",
        node.get("role") or "",
        node.get("title") or "",
        node.get("class") or "",
    ]).lower()
    
    for kw in keywords:
        if kw.lower() in combined:
            score += 5  # 其他属性中的匹配权重较低
    
    return score


def find_node_ref(
    snapshot: List[Dict[str, Any]],
    label: str,
    tag_whitelist: Optional[List[str]] = None,
) -> Optional[str]:
    """
    在快照中查找与标签匹配的元素
    
    优先匹配逻辑：
    1. 标签文本在相邻元素（prevSiblingText, rowLabelText）中，且当前元素是输入框
    2. 标签关键词在元素的 id、name 中
    3. 标签关键词在其他属性中
    
    只返回高置信度的匹配（分数 >= 20）
    """
    if not snapshot or not label:
        return None
    keywords = _tokenize_label(label)
    field_id = _extract_id_from_text(label)
    if field_id:
        keywords.append(field_id.lower())

    best_ref = None
    best_score = 0
    second_best_score = 0  # 用于检查是否有多个高分数匹配

    whitelist = [t.lower() for t in tag_whitelist] if tag_whitelist else None

    for node in snapshot:
        tag = (node.get("tag") or "").lower()
        if whitelist and tag not in whitelist:
            continue
        # 传递 label 文本用于上下文匹配
        score = _score_node(node, keywords, label_text=label)
        if score > best_score:
            second_best_score = best_score
            best_score = score
            best_ref = node.get("ref")
        elif score > second_best_score:
            second_best_score = score
    
    # 只返回高置信度的匹配（分数 >= 20，且最好比第二好的分数高至少 10 分）
    if best_score >= 20:
        # 如果第一和第二的分数太接近，可能匹配不准确，需要更严格的阈值
        if best_score - second_best_score >= 10 or second_best_score == 0:
            return best_ref
        # 如果分数接近，但第一名的分数很高（>= 50），仍然返回
        elif best_score >= 50:
            return best_ref
    
    return None


def validate_element_match(node: Dict[str, Any], label: str) -> bool:
    """
    验证元素是否真的匹配标签
    
    例如：如果查找"密码"，应该验证元素的 type 是 "password"
    """
    if not node or not label:
        return True  # 如果没有信息，默认通过验证
    
    label_lower = label.lower()
    node_id = (node.get("id") or "").lower()
    node_name = (node.get("name") or "").lower()
    node_type = (node.get("type") or "").lower()
    row_label = (node.get("rowLabelText") or "").lower()
    prev_sibling = (node.get("prevSiblingText") or "").lower()
    
    # 验证密码输入框
    if "密码" in label or "password" in label_lower or "pwd" in label_lower:
        # 密码输入框应该是 type="password"
        if node_type == "password":
            return True
        # 如果 id 或 name 明确包含 password/pwd，也接受
        if "password" in node_id or "pwd" in node_id or "password" in node_name or "pwd" in node_name:
            return True
        # 如果行标签或相邻元素包含"密码"，也接受
        if "密码" in row_label or "密码" in prev_sibling:
            return True
        # 否则可能是误匹配
        return False
    
    # 验证用户名输入框
    if "用户名" in label or "user" in label_lower or "login" in label_lower:
        # 用户名不应该是 password 类型
        if node_type == "password":
            return False
        # 如果 id 或 name 明确包含 user/login，接受
        if "user" in node_id or "login" in node_id or "user" in node_name or "login" in node_name:
            return True
        # 如果行标签或相邻元素包含"用户名"，接受
        if "用户名" in row_label or "用户名" in prev_sibling:
            return True
    
    # 其他字段，默认通过验证
    return True


def set_input_value_via_js(page: Page, ref: str, value: str) -> bool:
    try:
        page.evaluate(
            """(ref, value) => {
                const el = document.querySelector(`[data-mcp-ref="${ref}"]`);
                if (el) {
                    el.value = value;
                    el.dispatchEvent(new Event('input', { bubbles: true }));
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                }
            }""",
            ref,
            value,
        )
        return True
    except Exception:
        return False


def get_all_frames(page: Page) -> List[Any]:
    """获取页面中的所有 frame（包括主 frame 和所有 iframe）"""
    frames = [page.main_frame]
    # 递归获取所有子 frame
    def collect_frames(frame):
        for child in frame.child_frames:
            frames.append(child)
            collect_frames(child)
    collect_frames(page.main_frame)
    return frames


def find_element_in_frames(
    page: Page,
    selector: str,
    timeout: int = 5000,
    frame_context: Optional[Any] = None,
    debug: bool = False
) -> Optional[Tuple[Any, Any]]:
    """
    在所有 frame（主页面和所有 iframe）中查找元素
    
    Returns:
        (frame, locator) 如果找到，否则 None
    """
    frames = get_all_frames(page)
    
    for idx, frame in enumerate(frames):
        try:
            if selector.startswith("//"):
                locator = frame.locator(selector)
            else:
                locator = frame.locator(selector)
            
            # 检查元素是否存在
            count = locator.count()
            if count > 0:
                first = locator.first
                # 尝试检查可见性，但如果失败也继续尝试点击
                try:
                    if first.is_visible(timeout=1000):
                        if debug:
                            print(f"[DEBUG] 找到元素: {selector} 在 frame {idx}, 可见")
                        return (frame, locator)
                except Exception:
                    # 即使不可见也尝试返回（某些情况下元素可能被遮挡但仍然可点击）
                    if debug:
                        print(f"[DEBUG] 找到元素: {selector} 在 frame {idx}, 但不可见，仍尝试")
                    return (frame, locator)
        except Exception as e:
            if debug:
                print(f"[DEBUG] Frame {idx} 中查找失败: {e}")
            continue
    
    return None


def find_element_by_ref_in_frames(
    page: Page,
    ref: str,
    timeout: int = 5000
) -> Optional[Tuple[Any, Any]]:
    """
    在所有 frame 中通过 data-mcp-ref 查找元素
    
    Returns:
        (frame, locator) 如果找到，否则 None
    """
    frames = get_all_frames(page)
    
    for idx, frame in enumerate(frames):
        try:
            locator = frame.locator(f"[data-mcp-ref='{ref}']")
            count = locator.count()
            if count > 0:
                first = locator.first
                # 尝试检查可见性，但如果失败也继续尝试
                try:
                    if first.is_visible(timeout=1000):
                        return (frame, locator)
                except Exception:
                    # 即使不可见也尝试返回
                    return (frame, locator)
        except Exception:
            continue
    
    return None


def capture_dom_snapshot_with_frames(page: Page) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """
    捕获主页面和所有 iframe 中的 DOM 快照
    
    Returns:
        (snapshot_list, frame_map) - snapshot_list 包含所有元素，frame_map 记录每个元素所在的 frame
    """
    script = """
    (() => {
        const data = [];
        let counter = 1;
        const frameMap = {};
        
        function assignRef(el) {
            const ref = `mcp_${counter++}`;
            el.setAttribute('data-mcp-ref', ref);
            return ref;
        }
        
        function getSiblingText(el, direction) {
            // 获取相邻元素的文本（用于匹配标签）
            let sibling = direction === 'prev' ? el.previousElementSibling : el.nextElementSibling;
            if (sibling) {
                return (sibling.innerText || '').trim().substring(0, 100);
            }
            // 如果是表格单元格，尝试获取同一行的其他单元格
            if (el.tagName === 'TD' || el.closest('td')) {
                const td = el.tagName === 'TD' ? el : el.closest('td');
                const tr = td.parentElement;
                if (tr) {
                    const cells = Array.from(tr.children);
                    const idx = cells.indexOf(td);
                    if (direction === 'prev' && idx > 0) {
                        return (cells[idx - 1].innerText || '').trim().substring(0, 100);
                    } else if (direction === 'next' && idx < cells.length - 1) {
                        return (cells[idx + 1].innerText || '').trim().substring(0, 100);
                    }
                }
            }
            return '';
        }
        
        function walkFrame(frameDoc, frameId) {
            const walker = frameDoc.createTreeWalker(
                frameDoc.body || frameDoc.documentElement,
                NodeFilter.SHOW_ELEMENT
            );
            while (walker.nextNode()) {
                const el = walker.currentNode;
                const ref = assignRef(el);
                frameMap[ref] = frameId;
                const nodeData = {
                    ref,
                    tag: el.tagName || '',
                    id: el.id || '',
                    name: el.getAttribute('name') || '',
                    type: el.getAttribute('type') || '',
                    placeholder: el.getAttribute('placeholder') || '',
                    role: el.getAttribute('role') || '',
                    text: (el.innerText || '').trim().substring(0, 200)
                };
                
                // 添加相邻元素的文本（用于匹配标签）
                nodeData.prevSiblingText = getSiblingText(el, 'prev');
                nodeData.nextSiblingText = getSiblingText(el, 'next');
                
                // 添加父元素的文本（用于上下文匹配）
                const parent = el.parentElement;
                if (parent) {
                    nodeData.parentText = (parent.innerText || '').trim().substring(0, 200);
                    // 如果是表格，获取表头或标签单元格的文本
                    if (parent.tagName === 'TD' || parent.closest('td')) {
                        const td = parent.tagName === 'TD' ? parent : parent.closest('td');
                        const tr = td.parentElement;
                        if (tr) {
                            const cells = Array.from(tr.children);
                            const idx = cells.indexOf(td);
                            // 获取同一行第一个单元格的文本（通常是标签）
                            if (idx > 0 && cells[0]) {
                                nodeData.rowLabelText = (cells[0].innerText || '').trim().substring(0, 100);
                            }
                        }
                    }
                }
                
                // 添加图片 src 属性
                if (el.tagName === 'IMG') {
                    nodeData.src = el.getAttribute('src') || '';
                }
                // 添加 onclick 属性（用于查找可点击的 div）
                const onclick = el.getAttribute('onclick');
                if (onclick) {
                    nodeData.onclick = onclick.substring(0, 200);
                }
                // 添加 title 属性
                const title = el.getAttribute('title');
                if (title) {
                    nodeData.title = title;
                }
                // 添加 class 属性
                const className = el.getAttribute('class');
                if (className) {
                    nodeData.class = className;
                }
                data.push(nodeData);
            }
        }
        
        // 主文档
        walkFrame(document, 'main');
        
        // 所有 iframe
        const iframes = document.querySelectorAll('iframe');
        iframes.forEach((iframe, idx) => {
            try {
                const iframeDoc = iframe.contentDocument || iframe.contentWindow.document;
                if (iframeDoc) {
                    walkFrame(iframeDoc, `iframe_${idx}`);
                }
            } catch (e) {
                // 跨域 iframe 无法访问
            }
        });
        
        return { data, frameMap };
    })();
    """
    try:
        result = page.evaluate(script)
        return result.get("data", []), result.get("frameMap", {})
    except Exception:
        # 如果脚本失败，回退到只捕获主文档
        return capture_dom_snapshot(page), {}


def launch_browser(browser_type: str, headless: bool) -> Browser:
    options = {"headless": headless, "slow_mo": 100}
    browser_type = browser_type or "chromium"
    if browser_type == "chrome":
        browser_type = "chromium"
    pw = get_playwright()
    if browser_type == "chromium":
        return pw.chromium.launch(**options)
    if browser_type == "firefox":
        return pw.firefox.launch(**options)
    if browser_type == "webkit":
        return pw.webkit.launch(**options)
    return pw.chromium.launch(**options)


@dataclass
class BrowserSession:
    session_id: str
    browser: Browser
    page: Page
    headless: bool
    browser_type: str
    created_at: float
    last_used: float


class SessionManager:
    def __init__(self):
        self.sessions: Dict[str, BrowserSession] = {}
        self.lock = threading.Lock()

    def get(self, session_id: str) -> Optional[BrowserSession]:
        with self.lock:
            return self.sessions.get(session_id)

    def get_or_create(self, session_id: str, browser_type: str, headless: bool) -> BrowserSession:
        with self.lock:
            session = self.sessions.get(session_id)
        if session:
            if session.page.is_closed():
                try:
                    session.browser.close()
                except Exception:
                    pass
                browser = launch_browser(browser_type, headless)
                page = browser.new_page()
                page.set_default_timeout(300 * 1000)
                session.browser = browser
                session.page = page
            session.last_used = time.time()
            return session

        browser = launch_browser(browser_type, headless)
        page = browser.new_page()
        page.set_default_timeout(300 * 1000)
        new_session = BrowserSession(
            session_id=session_id,
            browser=browser,
            page=page,
            headless=headless,
            browser_type=browser_type,
            created_at=time.time(),
            last_used=time.time(),
        )
        with self.lock:
            self.sessions[session_id] = new_session
        return new_session

    def close_session(self, session_id: str) -> bool:
        with self.lock:
            session = self.sessions.pop(session_id, None)
        if not session:
            return False
        try:
            session.browser.close()
        except Exception:
            pass
        return True


session_manager = SessionManager()


app = FastAPI(title="Playwright MCP HTTP Gateway (Executor)")


class MCPRequest(BaseModel):
    """MCP 请求模型"""
    prompt: str
    timeout: Optional[int] = 300
    browser: Optional[str] = "chrome"
    headless: Optional[bool] = False  # 默认显示浏览器，方便调试
    session_id: Optional[str] = None


class MCPResponse(BaseModel):
    """MCP 响应模型"""
    status: str
    message: str
    execution_id: Optional[str] = None
    logs: Optional[list] = None
    error_details: Optional[Dict[str, Any]] = None
    timestamp: str


class CloseSessionRequest(BaseModel):
    session_id: str


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


def execute_step(page: Page, step: str, logs: List[str], snapshot: Optional[List[Dict[str, Any]]] = None) -> bool:
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
                
                # 准备备选选择器（仅在快照匹配失败时使用）
                default_selectors = [
                    f"label:has-text('{label}') + input",
                    f"input[placeholder*='{label}']",
                    f"//td[contains(text(), '{label}')]/following-sibling::td//input",
                ]
                selectors = _build_selector_candidates(label, INPUT_SELECTOR_HINTS, default_selectors)
                field_id = _extract_id_from_text(label)
                if field_id:
                    selectors.insert(0, f"#{field_id}")
                
                for attempt in range(MAX_STEP_RETRIES):
                    # 优先使用快照匹配（智能匹配，不依赖硬编码选择器）
                    if attempt == 0 and snapshot is not None:
                        current_snapshot = snapshot
                        frame_map = {}
                    else:
                        current_snapshot, frame_map = capture_dom_snapshot_with_frames(page)
                    
                    # 第一步：通过快照智能匹配（优先）
                    if current_snapshot:
                        ref = find_node_ref(current_snapshot, label, tag_whitelist=["input", "textarea"])
                        if ref:
                            # 验证匹配的元素是否真的正确
                            matched_node = next((n for n in current_snapshot if n.get("ref") == ref), None)
                            if matched_node and not validate_element_match(matched_node, label):
                                # 验证失败，跳过这个匹配，继续尝试其他方法
                                if attempt == 0:
                                    logs.append(f"⚠️  快照匹配到元素但验证失败，继续查找...")
                            else:
                                # 验证通过，尝试在所有 frame 中查找元素
                                result = find_element_by_ref_in_frames(page, ref)
                                if result:
                                    frame, locator = result
                                    try:
                                        locator.first.fill(value.strip(), timeout=5000)
                                        logs.append(f"✅ 已输入: {value.strip()} (通过快照匹配)")
                                        return True
                                    except Exception:
                                        # 如果 fill 失败，尝试使用 JS
                                        try:
                                            frame.evaluate(
                                                """(ref, value) => {
                                                    const el = document.querySelector(`[data-mcp-ref="${ref}"]`);
                                                    if (el) {
                                                        el.value = value;
                                                        el.dispatchEvent(new Event('input', { bubbles: true }));
                                                        el.dispatchEvent(new Event('change', { bubbles: true }));
                                                    }
                                                }""",
                                                ref,
                                                value.strip()
                                            )
                                            logs.append(f"✅ 已输入: {value.strip()} (通过快照匹配，JS)")
                                            return True
                                        except Exception:
                                            pass
                    
                    # 第二步：如果快照匹配失败，使用选择器作为备选
                    for selector in selectors:
                        result = find_element_in_frames(page, selector)
                        if result:
                            frame, locator = result
                            try:
                                locator.first.fill(value.strip(), timeout=5000)
                                logs.append(f"✅ 已输入: {value.strip()} (通过选择器: {selector[:50]})")
                                return True
                            except Exception as e:
                                if attempt == 0:
                                    logs.append(f"⚠️  找到元素但填充失败: {str(e)[:100]}")
                                continue
                    
                    if attempt < MAX_STEP_RETRIES - 1:
                        logs.append(f"⚠️  未找到输入框: {label}，5秒后重试（第{attempt + 1}次）")
                        page.wait_for_timeout(RETRY_INTERVAL_MS)
                
                # 调试信息：显示快照中的匹配情况
                if current_snapshot:
                    logs.append(f"🔍 调试信息: 快照中共有 {len(current_snapshot)} 个元素")
                    # 查找所有输入框，显示它们的上下文信息
                    input_nodes = [n for n in current_snapshot if n.get("tag", "").lower() in ["input", "textarea"]]
                    if input_nodes:
                        logs.append(f"   找到 {len(input_nodes)} 个输入框元素")
                        # 显示前3个最相关的输入框
                        keywords = _tokenize_label(label)
                        scored_nodes = [(n, _score_node(n, keywords, label_text=label)) for n in input_nodes]
                        scored_nodes.sort(key=lambda x: x[1], reverse=True)
                        for node, score in scored_nodes[:3]:
                            if score > 0:
                                node_id = node.get("id", "")[:30]
                                node_name = node.get("name", "")[:30]
                                row_label = node.get("rowLabelText", "")[:30]
                                logs.append(f"   候选: id={node_id}, name={node_name}, 行标签={row_label}, 分数={score}")
                
                logs.append(f"❌ 无法找到输入框: {label}")
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
                default_selectors = [
                    f"label:has-text('{label}') + select",
                    f"select[name*='{label}']",
                    f"select[id*='{label}']",
                    f"//label[contains(text(), '{label}')]/following-sibling::select[1]",
                    # 通过表格单元格查找（适用于表单在表格中的情况）
                    f"//td[contains(text(), '{label}')]/following-sibling::td//select",
                    f"//td[contains(text(), '{label}')]/following-sibling::td[1]//select",
                    f"//td[@class='iscap' and contains(text(), '{label}')]/following-sibling::td//select",
                    f"//td[@class='iscap' and contains(text(), '{label}')]/following-sibling::td[1]//select"
                ]
                selectors = _build_selector_candidates(label, DROPDOWN_SELECTOR_HINTS, default_selectors)
                selector_id = _extract_id_from_text(label)
                if selector_id:
                    selectors.insert(0, f"select#{selector_id}")
                
                # 对于特定字段，添加更精确的选择器
                if "支付方式" in label or "pay_type" in label.lower() or "payType" in label:
                    selectors.insert(0, "select[id*='pay_type']")
                    selectors.insert(0, "select[name*='pay_type']")
                for attempt in range(MAX_STEP_RETRIES):
                    # 使用支持 iframe 的快照捕获
                    if attempt == 0 and snapshot is not None:
                        current_snapshot = snapshot
                        frame_map = {}
                    else:
                        current_snapshot, frame_map = capture_dom_snapshot_with_frames(page)
                    
                    # 第一步：通过快照智能匹配（优先）
                    if current_snapshot:
                        ref = find_node_ref(current_snapshot, label, tag_whitelist=["select"])
                        if ref:
                            # 尝试在所有 frame 中查找元素
                            result = find_element_by_ref_in_frames(page, ref)
                            if result:
                                frame, locator = result
                                try:
                                    # 尝试通过文本选择（优先）
                                    try:
                                        locator.first.select_option(label=value, timeout=5000)
                                        logs.append(f"✅ 已选择: {value} (通过快照匹配，文本选择)")
                                        page.wait_for_timeout(500)
                                        return True
                                    except Exception:
                                        # 如果文本选择失败，尝试通过值选择
                                        locator.first.select_option(value, timeout=5000)
                                        logs.append(f"✅ 已选择: {value} (通过快照匹配，值选择)")
                                        page.wait_for_timeout(500)
                                        return True
                                except Exception as e:
                                    if attempt == 0:
                                        logs.append(f"⚠️  找到下拉框但选择失败: {str(e)[:100]}")
                                    pass
                    
                    # 第二步：如果快照匹配失败，使用选择器作为备选
                    for selector in selectors:
                        result = find_element_in_frames(page, selector)
                        if result:
                            frame, locator = result
                            try:
                                # 尝试通过文本选择（优先）
                                try:
                                    locator.first.select_option(label=value, timeout=5000)
                                    logs.append(f"✅ 已选择: {value} (通过选择器，文本选择)")
                                    page.wait_for_timeout(500)
                                    return True
                                except Exception:
                                    # 如果文本选择失败，尝试通过值选择
                                    locator.first.select_option(value, timeout=5000)
                                    logs.append(f"✅ 已选择: {value} (通过选择器，值选择)")
                                    page.wait_for_timeout(500)
                                    return True
                            except Exception as e:
                                if attempt == 0:
                                    logs.append(f"⚠️  找到下拉框但选择失败: {str(e)[:100]}")
                                continue
                    
                    if attempt < MAX_STEP_RETRIES - 1:
                        logs.append(f"⚠️  未找到下拉框: {label}，5秒后重试（第{attempt + 1}次）")
                        page.wait_for_timeout(RETRY_INTERVAL_MS)
                
                # 调试信息：显示快照中的匹配情况
                if current_snapshot:
                    logs.append(f"🔍 调试信息: 快照中共有 {len(current_snapshot)} 个元素")
                    # 查找所有下拉框，显示它们的上下文信息
                    select_nodes = [n for n in current_snapshot if n.get("tag", "").lower() == "select"]
                    if select_nodes:
                        logs.append(f"   找到 {len(select_nodes)} 个下拉框元素")
                        # 显示前3个最相关的下拉框
                        keywords = _tokenize_label(label)
                        scored_nodes = [(n, _score_node(n, keywords, label_text=label)) for n in select_nodes]
                        scored_nodes.sort(key=lambda x: x[1], reverse=True)
                        for node, score in scored_nodes[:3]:
                            if score > 0:
                                node_id = node.get("id", "")[:50]
                                node_name = node.get("name", "")[:50]
                                row_label = node.get("rowLabelText", "")[:30]
                                logs.append(f"   候选: id={node_id}, name={node_name}, 行标签={row_label}, 分数={score}")
                
                logs.append(f"❌ 无法找到下拉框: {label}")
                return False
            return False
        
        # 点击 radio button
        elif "radio button" in step and "点击" in step:
            match = re.search(r'点击(.+?)radio button', step)
            if match:
                radio_text = match.group(1).strip()
                logs.append(f"正在点击 radio button: {radio_text}")
                
                # 准备选择器（优先通过文本查找，因为 radio button 通常在 li 或 span 中）
                default_selectors = [
                    # 通过文本查找：li 中包含文本，然后找其中的 radio button
                    f"//li[.//span[contains(text(), '{radio_text}')]]//input[@type='radio']",
                    f"//li[contains(text(), '{radio_text}')]//input[@type='radio']",
                    # 通过文本查找：span 中包含文本，然后找同级的 radio button
                    f"//span[contains(text(), '{radio_text}')]/preceding-sibling::input[@type='radio']",
                    f"//span[contains(text(), '{radio_text}')]/../input[@type='radio']",
                    # 通过 value 查找
                    f"input[type='radio'][value*='{radio_text}']",
                    f"input[type='radio'][name*='filter_bcode'][value*='{radio_text}']",
                ]
                
                for attempt in range(MAX_STEP_RETRIES):
                    # 使用支持 iframe 的快照捕获
                    if attempt == 0 and snapshot is not None:
                        current_snapshot = snapshot
                        frame_map = {}
                    else:
                        current_snapshot, frame_map = capture_dom_snapshot_with_frames(page)
                    
                    # 第一步：通过快照智能匹配（优先）
                    if current_snapshot:
                        # 查找 radio button，但也要考虑通过相邻文本匹配
                        ref = find_node_ref(current_snapshot, radio_text, tag_whitelist=["input"])
                        if ref:
                            result = find_element_by_ref_in_frames(page, ref)
                            if result:
                                frame, locator = result
                                try:
                                    # 检查是否是 radio button
                                    element_type = locator.first.get_attribute("type")
                                    if element_type == "radio":
                                        locator.first.click(timeout=5000)
                                        logs.append(f"✅ 已点击 radio button: {radio_text} (通过快照匹配)")
                                        page.wait_for_timeout(500)
                                        return True
                                except Exception as e:
                                    if attempt == 0:
                                        logs.append(f"⚠️  找到元素但点击失败: {str(e)[:100]}")
                                    pass
                        
                        # 如果直接匹配失败，尝试查找包含文本的 span，然后找其父级或相邻的 radio button
                        # 在快照中查找包含文本的节点（可能是 span 或 li）
                        for node in current_snapshot:
                            node_text = (node.get("text") or "").lower()
                            node_tag = (node.get("tag") or "").lower()
                            if radio_text.lower() in node_text and node_tag in ["span", "li"]:
                                # 找到包含文本的节点，尝试通过其 ref 查找同级的 radio button
                                node_ref = node.get("ref")
                                if node_ref:
                                    # 尝试查找父级或相邻的 radio button
                                    try:
                                        # 方法1：查找父级 li 中的 radio button
                                        parent_radio_sel = f"li:has(span[data-mcp-ref='{node_ref}']) input[type='radio']"
                                        result = find_element_in_frames(page, parent_radio_sel)
                                        if result:
                                            frame, locator = result
                                            locator.first.click(timeout=5000)
                                            logs.append(f"✅ 已点击 radio button: {radio_text} (通过文本节点匹配)")
                                            page.wait_for_timeout(500)
                                            return True
                                    except Exception:
                                        pass
                    
                    # 第二步：使用选择器查找
                    for selector in default_selectors:
                        result = find_element_in_frames(page, selector)
                        if result:
                            frame, locator = result
                            try:
                                locator.first.click(timeout=5000)
                                logs.append(f"✅ 已点击 radio button: {radio_text} (通过选择器)")
                                page.wait_for_timeout(500)
                                return True
                            except Exception as e:
                                if attempt == 0:
                                    logs.append(f"⚠️  找到元素但点击失败: {str(e)[:100]}")
                                continue
                    
                    if attempt < MAX_STEP_RETRIES - 1:
                        logs.append(f"⚠️  未找到 radio button: {radio_text}，5秒后重试（第{attempt + 1}次）")
                        page.wait_for_timeout(RETRY_INTERVAL_MS)
                
                logs.append(f"❌ 无法找到 radio button: {radio_text}")
                return False
            return False
        
        # 点击按钮
        elif "按钮" in step and "点击" in step:
            # 检查是否是"等待...出现后，点击"格式
            wait_match = re.search(r'等待(.+?)按钮出现后，点击', step)
            if wait_match:
                button_text = wait_match.group(1)
                logs.append(f"等待按钮出现: {button_text}")
                # 先等待元素出现（使用智能匹配）
                wait_success = False
                wait_timeout = 30000  # 30秒超时
                wait_start = time.time()
                
                while time.time() - wait_start < wait_timeout / 1000:
                    # 方法1：使用快照匹配（最智能）
                    try:
                        current_snapshot, _ = capture_dom_snapshot_with_frames(page)
                        if current_snapshot:
                            ref = find_node_ref(current_snapshot, button_text, tag_whitelist=["button", "a", "input", "div"])
                            if ref:
                                result = find_element_by_ref_in_frames(page, ref)
                                if result:
                                    frame, locator = result
                                    try:
                                        if locator.first.is_visible(timeout=1000):
                                            wait_success = True
                                            logs.append(f"✅ 按钮已出现: {button_text} (通过快照匹配)")
                                            break
                                    except Exception:
                                        pass
                    except Exception:
                        pass
                    
                    # 方法2：使用选择器查找
                    if not wait_success:
                        frames = get_all_frames(page)
                        for frame in frames:
                            try:
                                # 尝试多种选择器
                                test_selectors = [
                                    f"button:has-text('{button_text}')",
                                    f"a:has-text('{button_text}')",
                                    f"div:has-text('{button_text}')",
                                    f"//button[contains(text(), '{button_text}')]",
                                    f"//a[contains(text(), '{button_text}')]",
                                ]
                                for sel in test_selectors:
                                    try:
                                        locator = frame.locator(sel)
                                        if locator.count() > 0:
                                            if locator.first.is_visible(timeout=1000):
                                                wait_success = True
                                                logs.append(f"✅ 按钮已出现: {button_text} (通过选择器)")
                                                break
                                    except Exception:
                                        continue
                                if wait_success:
                                    break
                            except Exception:
                                continue
                    
                    if wait_success:
                        break
                    
                    page.wait_for_timeout(500)  # 等待500ms后重试
                
                if not wait_success:
                    logs.append(f"⚠️  等待超时，按钮未出现: {button_text}，继续尝试点击...")
                    # 继续尝试点击，可能元素已经存在但不可见
                
                # 继续执行点击逻辑
            else:
                # 普通点击格式
                match = re.search(r'点击(.+?)按钮', step)
                if not match:
                    return False
                button_text = match.group(1)
            
            logs.append(f"正在点击按钮: {button_text}")
            default_selectors = [
                f"button:has-text('{button_text}')",
                f"a:has-text('{button_text}')",
                f"div:has-text('{button_text}')",
                f"div[onclick]:has-text('{button_text}')",
                f"input[type='button'][value='{button_text}']",
                f"input[type='submit'][value='{button_text}']",
                f"//button[contains(text(), '{button_text}')]",
                f"//a[contains(text(), '{button_text}')]",
                f"//div[contains(text(), '{button_text}')]",
                f"//div[@onclick and contains(text(), '{button_text}')]"
            ]
            selectors = _deduplicate_selectors(BUTTON_SELECTOR_HINTS.get(button_text, []) + default_selectors)
            button_id = _extract_id_from_text(button_text)
            if button_id:
                selectors.insert(0, f"#{button_id}")
            
            # 特殊处理：如果按钮文本包含特定关键词，添加通过图片查找的选择器
            if "预约" in button_text or "报账" in button_text:
                # 尝试通过图片路径查找父元素
                img_based_selectors = [
                    "div:has(img[src*='wsyy'])",
                    "div:has(img[src*='预约'])",
                    "div:has(img[src*='报账'])",
                    "div.syslink:has(img)",
                    "div[onclick]:has(img)"
                ]
                selectors = img_based_selectors + selectors
            
            for attempt in range(MAX_STEP_RETRIES):
                # 使用支持 iframe 的快照捕获
                if attempt == 0 and snapshot is not None:
                    current_snapshot = snapshot
                    frame_map = {}
                else:
                    current_snapshot, frame_map = capture_dom_snapshot_with_frames(page)
                
                # 优先尝试通过选择器直接在所有 frame 中查找（不依赖快照）
                # 这对于 iframe 中的元素更可靠
                if "预约" in button_text or "报账" in button_text:
                    # 特殊处理：网上预约报账按钮
                    special_selectors = [
                        "div.syslink:has(img[src*='wsyy.jpg'])",
                        "div[onclick*='WF_YB6']",
                        "div[onclick*='navToPrj']:has(img[src*='wsyy'])",
                        "div.syslink[title*='点击进入']:has(img[src*='wsyy'])",
                        "div:has(img[src*='wsyy.jpg'])",
                        "div:has(img[src*='wsyy'])",
                        "div.syslink:has(img)",
                        "div[onclick]:has(img[src*='wsyy'])"
                    ]
                    for special_sel in special_selectors:
                        result = find_element_in_frames(page, special_sel, debug=(attempt == 0))
                        if result:
                            frame, locator = result
                            try:
                                element = locator.first
                                element.wait_for(state="visible", timeout=5000)
                                element.click(timeout=5000)
                                logs.append(f"✅ 已点击: {button_text} (通过特殊选择器，在 iframe 中)")
                                # 根据按钮类型决定等待时间
                                if "申请" in button_text or "提交" in button_text or "下一步" in button_text:
                                    logs.append(f"⏳ 等待页面响应...")
                                    page.wait_for_timeout(2000)  # 先等待2秒
                                    try:
                                        page.wait_for_load_state("networkidle", timeout=15000)
                                        logs.append(f"✅ 页面加载完成")
                                    except Exception:
                                        page.wait_for_timeout(2000)
                                        logs.append(f"⚠️  网络空闲超时，继续等待2秒")
                                else:
                                    page.wait_for_timeout(2000)
                                return True
                            except Exception as e:
                                logs.append(f"⚠️  找到元素但点击失败: {str(e)}")
                                continue
                
                # 尝试使用预定义的选择器在所有 frame 中查找
                for selector in selectors:
                    result = find_element_in_frames(page, selector, debug=(attempt == 0))
                    if result:
                        frame, locator = result
                        try:
                            element = locator.first
                            element.wait_for(state="visible", timeout=5000)
                            element.click(timeout=5000)
                            logs.append(f"✅ 已点击: {button_text} (在 iframe 中)")
                            # 根据按钮类型决定等待时间
                            if "申请" in button_text or "提交" in button_text or "下一步" in button_text:
                                logs.append(f"⏳ 等待页面响应...")
                                page.wait_for_timeout(3000)
                                try:
                                    page.wait_for_load_state("networkidle", timeout=10000)
                                except Exception:
                                    pass
                            else:
                                page.wait_for_timeout(2000)
                            return True
                        except Exception as e:
                            logs.append(f"⚠️  找到元素但点击失败: {str(e)}")
                            continue
                
                # 如果快照可用，尝试通过快照查找
                if current_snapshot:
                    # 扩展标签白名单，包含 div
                    ref = find_node_ref(current_snapshot, button_text, tag_whitelist=["button", "a", "input", "div"])
                    if ref:
                        # 尝试在所有 frame 中查找元素
                        result = find_element_by_ref_in_frames(page, ref)
                        if result:
                            frame, locator = result
                            try:
                                # 点击前确保元素可见且可点击
                                element = locator.first
                                element.wait_for(state="visible", timeout=5000)
                                
                                # 记录当前URL（用于验证页面是否跳转）
                                current_url = page.url
                                
                                # 执行点击
                                element.click(timeout=5000)
                                logs.append(f"✅ 已点击: {button_text} (通过快照 ref，在 iframe 中)")
                                
                                # 点击后等待，根据按钮类型决定等待时间
                                if "申请" in button_text or "提交" in button_text or "下一步" in button_text:
                                    # 这些按钮通常会触发页面跳转或加载新内容，需要等待更长时间
                                    logs.append(f"⏳ 等待页面响应...")
                                    page.wait_for_timeout(2000)  # 先等待2秒
                                    # 等待网络空闲或页面加载完成
                                    try:
                                        page.wait_for_load_state("networkidle", timeout=15000)
                                        logs.append(f"✅ 页面加载完成")
                                    except Exception:
                                        # 如果网络空闲超时，至少等待一段时间确保内容加载
                                        page.wait_for_timeout(2000)
                                        logs.append(f"⚠️  网络空闲超时，继续等待2秒")
                                else:
                                    page.wait_for_timeout(2000)  # 普通按钮等待2秒
                                
                                return True
                            except Exception as e:
                                logs.append(f"⚠️  点击失败: {str(e)[:100]}")
                                pass
                    
                    # 如果通过文本找不到，尝试通过图片查找
                    if "预约" in button_text or "报账" in button_text:
                        # 通过图片 src 查找
                        img_ref = find_node_ref(current_snapshot, "wsyy", tag_whitelist=["img"])
                        if img_ref:
                            # 尝试找到图片的父 div
                            result = find_element_in_frames(page, f"div:has(img[data-mcp-ref='{img_ref}'])")
                            if result:
                                frame, locator = result
                                try:
                                    element = locator.first
                                    element.wait_for(state="visible", timeout=5000)
                                    element.click(timeout=5000)
                                    logs.append(f"✅ 已点击: {button_text} (通过图片 ref 定位)")
                                    # 根据按钮类型决定等待时间
                                    if "申请" in button_text or "提交" in button_text or "下一步" in button_text:
                                        logs.append(f"⏳ 等待页面响应...")
                                        page.wait_for_timeout(3000)
                                        try:
                                            page.wait_for_load_state("networkidle", timeout=10000)
                                        except Exception:
                                            pass
                                    else:
                                        page.wait_for_timeout(2000)
                                    return True
                                except Exception as e:
                                    logs.append(f"⚠️  点击失败: {str(e)[:100]}")
                                    pass
                
                if attempt < MAX_STEP_RETRIES - 1:
                    logs.append(f"⚠️  未找到按钮: {button_text}，5秒后重试（第{attempt + 1}次）")
                    page.wait_for_timeout(RETRY_INTERVAL_MS)
            
            # 最后一次尝试：列出所有 frame 信息用于调试
            frames = get_all_frames(page)
            logs.append(f"🔍 调试信息: 共找到 {len(frames)} 个 frame")
            for idx, frame in enumerate(frames):
                try:
                    # 尝试查找任何包含 wsyy 的元素
                    test_locator = frame.locator("img[src*='wsyy'], div:has(img[src*='wsyy'])")
                    count = test_locator.count()
                    if count > 0:
                        logs.append(f"   Frame {idx}: 找到 {count} 个包含 'wsyy' 的元素")
                except Exception:
                    pass
            
            logs.append(f"❌ 无法找到按钮: {button_text}")
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
                for attempt in range(MAX_STEP_RETRIES):
                    # 使用支持 iframe 的快照捕获
                    if attempt == 0 and snapshot is not None:
                        current_snapshot = snapshot
                        frame_map = {}
                    else:
                        current_snapshot, frame_map = capture_dom_snapshot_with_frames(page)
                    
                    if current_snapshot:
                        ref = find_node_ref(current_snapshot, label, tag_whitelist=["input", "textarea"])
                        if ref:
                            # 尝试在所有 frame 中查找元素
                            result = find_element_by_ref_in_frames(page, ref)
                            if result:
                                frame, locator = result
                                try:
                                    locator.first.fill(value.strip(), timeout=5000)
                                    logs.append(f"✅ 已填写: {value.strip()} (在 iframe 中)")
                                    return True
                                except Exception:
                                    try:
                                        frame.evaluate(
                                            """(ref, value) => {
                                                const el = document.querySelector(`[data-mcp-ref="${ref}"]`);
                                                if (el) {
                                                    el.value = value;
                                                    el.dispatchEvent(new Event('input', { bubbles: true }));
                                                    el.dispatchEvent(new Event('change', { bubbles: true }));
                                                }
                                            }""",
                                            ref,
                                            value.strip()
                                        )
                                        logs.append(f"✅ 已填写: {value.strip()} (JS, 在 iframe 中)")
                                        return True
                                    except Exception:
                                        pass
                    
                    # 尝试使用选择器在所有 frame 中查找
                    for selector in selectors:
                        result = find_element_in_frames(page, selector)
                        if result:
                            frame, locator = result
                            try:
                                locator.first.fill(value.strip(), timeout=5000)
                                logs.append(f"✅ 已填写: {value.strip()} (在 iframe 中)")
                                return True
                            except Exception:
                                continue
                    
                    if attempt < MAX_STEP_RETRIES - 1:
                        logs.append(f"⚠️  未找到输入框: {label}，5秒后重试（第{attempt + 1}次）")
                        page.wait_for_timeout(RETRY_INTERVAL_MS)
                logs.append(f"❌ 无法找到输入框: {label}")
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
                for attempt in range(MAX_STEP_RETRIES):
                    # 使用支持 iframe 的快照捕获
                    if attempt == 0 and snapshot is not None:
                        current_snapshot = snapshot
                        frame_map = {}
                    else:
                        current_snapshot, frame_map = capture_dom_snapshot_with_frames(page)
                    
                    if current_snapshot:
                        ref = find_node_ref(current_snapshot, label, tag_whitelist=["input"])
                        if ref:
                            # 尝试在所有 frame 中查找元素
                            result = find_element_by_ref_in_frames(page, ref)
                            if result:
                                frame, locator = result
                                try:
                                    locator.first.fill(date.strip(), timeout=5000)
                                    logs.append(f"✅ 已选择日期: {date.strip()} (在 iframe 中)")
                                    return True
                                except Exception:
                                    try:
                                        frame.evaluate(
                                            """(ref, value) => {
                                                const el = document.querySelector(`[data-mcp-ref="${ref}"]`);
                                                if (el) {
                                                    el.value = value;
                                                    el.dispatchEvent(new Event('input', { bubbles: true }));
                                                    el.dispatchEvent(new Event('change', { bubbles: true }));
                                                }
                                            }""",
                                            ref,
                                            date.strip()
                                        )
                                        logs.append(f"✅ 已选择日期: {date.strip()} (JS, 在 iframe 中)")
                                        return True
                                    except Exception:
                                        pass
                    
                    # 尝试使用选择器在所有 frame 中查找
                    for selector in selectors:
                        result = find_element_in_frames(page, selector)
                        if result:
                            frame, locator = result
                            try:
                                locator.first.fill(date.strip(), timeout=5000)
                                logs.append(f"✅ 已选择日期: {date.strip()} (在 iframe 中)")
                                return True
                            except Exception:
                                continue
                    
                    if attempt < MAX_STEP_RETRIES - 1:
                        logs.append(f"⚠️  未找到日期选择器: {label}，5秒后重试（第{attempt + 1}次）")
                        page.wait_for_timeout(RETRY_INTERVAL_MS)
                logs.append(f"❌ 无法找到日期选择器: {label}")
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
                for attempt in range(MAX_STEP_RETRIES):
                    # 使用支持 iframe 的快照捕获
                    if attempt == 0 and snapshot is not None:
                        current_snapshot = snapshot
                        frame_map = {}
                    else:
                        current_snapshot, frame_map = capture_dom_snapshot_with_frames(page)
                    
                    if current_snapshot:
                        ref = find_node_ref(current_snapshot, "银行卡号", tag_whitelist=["input"])
                        if ref:
                            # 尝试在所有 frame 中查找元素
                            result = find_element_by_ref_in_frames(page, ref)
                            if result:
                                frame, locator = result
                                try:
                                    locator.first.fill(tail, timeout=5000)
                                    logs.append(f"✅ 已输入银行卡号尾号: {tail} (在 iframe 中)")
                                    return True
                                except Exception:
                                    try:
                                        frame.evaluate(
                                            """(ref, value) => {
                                                const el = document.querySelector(`[data-mcp-ref="${ref}"]`);
                                                if (el) {
                                                    el.value = value;
                                                    el.dispatchEvent(new Event('input', { bubbles: true }));
                                                    el.dispatchEvent(new Event('change', { bubbles: true }));
                                                }
                                            }""",
                                            ref,
                                            tail
                                        )
                                        logs.append(f"✅ 已输入银行卡号尾号: {tail} (JS, 在 iframe 中)")
                                        return True
                                    except Exception:
                                        pass
                    
                    # 尝试使用选择器在所有 frame 中查找
                    for selector in selectors:
                        result = find_element_in_frames(page, selector)
                        if result:
                            frame, locator = result
                            try:
                                locator.first.fill(tail, timeout=5000)
                                logs.append(f"✅ 已输入银行卡号尾号: {tail} (在 iframe 中)")
                                return True
                            except Exception:
                                continue
                    
                    if attempt < MAX_STEP_RETRIES - 1:
                        logs.append(f"⚠️  未找到银行卡号输入框，5秒后重试（第{attempt + 1}次）")
                        page.wait_for_timeout(RETRY_INTERVAL_MS)
                logs.append("❌ 无法找到银行卡号输入框")
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
                    img_selectors = IMAGE_SELECTOR_HINTS.get("验证码", []) + [
                        "img[src*='captcha']",
                        "img[src*='code']",
                        "img[alt*='验证码']",
                        "img[id*='captcha']",
                        "img[id*='code']",
                        "//img[contains(@src, 'captcha')]",
                        "//img[contains(@src, 'code')]"
                    ]
                    id_hint = _extract_id_from_text(step)
                    if id_hint:
                        img_selectors.insert(0, f"#{id_hint}")
                    for attempt in range(MAX_STEP_RETRIES):
                        # 使用支持 iframe 的快照捕获
                        if attempt == 0 and snapshot is not None:
                            current_snapshot = snapshot
                            frame_map = {}
                        else:
                            current_snapshot, frame_map = capture_dom_snapshot_with_frames(page)
                        
                        if current_snapshot:
                            ref = find_node_ref(current_snapshot, "验证码", tag_whitelist=["img", "canvas"])
                            if ref:
                                # 尝试在所有 frame 中查找元素
                                result = find_element_by_ref_in_frames(page, ref)
                                if result:
                                    frame, locator = result
                                    try:
                                        save_path = os.path.join(save_dir, filename)
                                        os.makedirs(save_dir, exist_ok=True)
                                        locator.first.screenshot(path=save_path)
                                        logs.append(f"✅ 验证码图片已保存: {save_path} (在 iframe 中)")
                                        return True
                                    except Exception:
                                        pass
                        
                        # 尝试使用选择器在所有 frame 中查找
                        for selector in img_selectors:
                            result = find_element_in_frames(page, selector)
                            if result:
                                frame, locator = result
                                try:
                                    save_path = os.path.join(save_dir, filename)
                                    os.makedirs(save_dir, exist_ok=True)
                                    locator.first.screenshot(path=save_path)
                                    logs.append(f"✅ 验证码图片已保存: {save_path} (在 iframe 中)")
                                    return True
                                except Exception:
                                    continue
                        
                        if attempt < MAX_STEP_RETRIES - 1:
                            logs.append(f"⚠️  未找到验证码图片，5秒后重试（第{attempt + 1}次）")
                            page.wait_for_timeout(RETRY_INTERVAL_MS)
                    logs.append(f"❌ 未找到验证码图片")
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
            # 尝试从步骤中提取等待时间（秒）
            wait_seconds = None
            # 匹配格式：等待 X 秒、等待X秒、等待 X秒、等待 X 秒钟等
            match = re.search(r'等待\s*(\d+)\s*秒', step)
            if match:
                wait_seconds = int(match.group(1))
            
            if wait_seconds:
                wait_ms = wait_seconds * 1000
                logs.append(f"等待: {wait_seconds} 秒")
                page.wait_for_timeout(wait_ms)
                logs.append(f"✅ 等待完成")
            elif "页面响应" in step or "页面跳转" in step:
                logs.append(f"等待页面响应...")
                page.wait_for_timeout(3000)
                logs.append(f"✅ 等待完成")
            else:
                logs.append(f"等待: {step}")
                page.wait_for_timeout(2000)
                logs.append(f"✅ 等待完成")
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


def execute_playwright_prompt(
    prompt: str,
    headless: bool = False,
    browser_type: str = "chromium",
    timeout: int = 300,
    session: Optional[BrowserSession] = None,
    persist_session: bool = False,
) -> Dict[str, Any]:
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
    logs = LiveLog()
    steps = parse_mcp_prompt(prompt)
    
    logs.append(f"🚀 开始执行，共 {len(steps)} 个步骤")
    
    browser: Optional[Browser] = None
    page: Optional[Page] = None
    created_new_browser = False

    try:
        if session:
            browser = session.browser
            page = session.page
            logs.append(f"♻️  复用会话 {session.session_id}")
        else:
            browser = launch_browser(browser_type, headless)
            page = browser.new_page()
            created_new_browser = True
            logs.append(f"✅ 浏览器已启动: {browser_type} (headless={headless})")
            logs.append("✅ 新页面已创建")

        if page is None:
            raise RuntimeError("页面初始化失败")

        page.set_default_timeout(timeout * 1000)

        success_count = 0
        failed_steps: List[Tuple[int, str]] = []

        for i, step in enumerate(steps, 1):
            logs.append(f"\n--- 步骤 {i}/{len(steps)}: {step[:50]}... ---")
            # 使用支持 iframe 的快照捕获
            snapshot, _ = capture_dom_snapshot_with_frames(page)
            if execute_step(page, step, logs, snapshot):
                success_count += 1
            else:
                failed_steps.append((i, step))

        # 如果复用会话，不等待和关闭浏览器
        if created_new_browser and not headless and not session:
            logs.append("\n⏳ 等待 3 秒后关闭浏览器...")
            try:
                page.wait_for_timeout(3000)
            except Exception:
                pass

        result = {
            "execution_id": execution_id,
            "logs": logs,
            "total_steps": len(steps),
            "success_steps": success_count,
            "failed_steps": failed_steps,
        }

        if len(failed_steps) == 0:
            result.update({
                "status": "success",
                "message": f"所有步骤执行成功（{success_count}/{len(steps)}）",
            })
        else:
            result.update({
                "status": "partial",
                "message": f"部分步骤执行成功（{success_count}/{len(steps)}）",
            })
        return result

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
    finally:
        # 只有在明确要求关闭会话时才关闭，否则保持会话打开
        if session:
            session.last_used = time.time()
            # 不再自动关闭会话，保持会话打开以便后续使用
            # if not persist_session:
            #     session_manager.close_session(session.session_id)
        elif created_new_browser and browser and not persist_session:
            # 只有在没有会话且明确不保持会话时才关闭浏览器
            try:
                browser.close()
                logs.append("\n✅ 浏览器已关闭")
            except Exception:
                pass


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
            "health": "/health",
            "close_session": "/mcp/close-session"
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
        browser_type = request.browser or "chrome"
        headless = request.headless if request.headless is not None else False
        timeout = request.timeout or 300
        
        # 如果没有提供 session_id，自动生成一个（用于保持会话）
        if request.session_id and request.session_id.strip():
            session_id = request.session_id.strip()
        else:
            # 自动生成 session_id，使用时间戳和随机数
            import random
            session_id = f"auto_{int(time.time())}_{random.randint(1000, 9999)}"
        
        persist_session = True  # 默认保持会话打开
        
        logger.info(
            "收到执行请求 headless=%s browser=%s timeout=%s session_id=%s",
            headless,
            browser_type,
            timeout,
            session_id,
        )

        loop = asyncio.get_running_loop()

        def run_task():
            # 总是获取或创建会话（因为 persist_session 总是 True）
            session_obj = session_manager.get_or_create(session_id, browser_type, headless)
            return execute_playwright_prompt(
                prompt=request.prompt,
                headless=headless,
                browser_type=browser_type,
                timeout=timeout,
                session=session_obj,
                persist_session=persist_session,
            )

        result = await loop.run_in_executor(PLAYWRIGHT_EXECUTOR, run_task)
        result["timestamp"] = datetime.now().isoformat()
        # 始终返回 session_id，以便后续请求可以复用会话
        result["session_id"] = session_id
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


@app.post("/mcp/close-session")
async def close_session_endpoint(request: CloseSessionRequest):
    loop = asyncio.get_running_loop()

    def task():
        return session_manager.close_session(request.session_id)

    closed = await loop.run_in_executor(PLAYWRIGHT_EXECUTOR, task)
    if closed:
        return {
            "status": "success",
            "message": f"会话 {request.session_id} 已关闭",
            "timestamp": datetime.now().isoformat()
        }
    return JSONResponse(
        status_code=404,
        content={
            "status": "not_found",
            "message": f"未找到会话 {request.session_id}",
            "timestamp": datetime.now().isoformat()
        }
    )


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

