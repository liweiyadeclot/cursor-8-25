#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
调试模型响应
"""

import requests
import json

def debug_model_response():
    """
    调试模型响应，查看原始输出
    """
    print("=== 调试模型响应 ===")
    
    # 测试输入
    test_input = "张三，学工号2021001，技术部，差旅费500元"
    
    # 构建提示词
    prompt = f"""你是一个专业的财务报销信息提取助手。请从用户输入的自然语言中提取报销相关信息，并按照以下7个阶段进行分类整理。

提取阶段说明：
1. 登录阶段：用户名、密码
2. 项目信息：项目号、附件张数、支付方式
3. 报销科目：科目类型、金额
4. 报销人员：学工号、银行卡信息
5. 预约报销：报销时间、地点
6. 差旅信息：姓名、人员类型、出差地点
7. 劳务信息：劳务费类型、发放事由

请仔细分析用户输入，提取出相关信息。如果某个阶段没有相关信息，请在该阶段中填写null。

重要：请严格按照以下JSON格式输出，使用英文键名，不要包含任何其他文字说明。

{{
  "login": {{
    "username": "用户名",
    "password": "密码"
  }},
  "project": {{
    "projectNumber": "项目号",
    "attachmentCount": "附件张数",
    "paymentMethod": "支付方式"
  }},
  "expense": {{
    "category": "科目类型",
    "amount": "金额"
  }},
  "personnel": {{
    "studentId": "学工号",
    "bankCard": "银行卡信息"
  }},
  "appointment": {{
    "date": "报销时间",
    "location": "地点"
  }},
  "travel": {{
    "name": "姓名",
    "personnelType": "人员类型",
    "destination": "出差地点"
  }},
  "labor": {{
    "laborType": "劳务费类型",
    "reason": "发放事由"
  }}
}}

用户输入：{test_input}

JSON输出："""
    
    try:
        print("发送请求到模型...")
        response = requests.post(
            "http://localhost:11434/api/generate",
            json={
                "model": "qwen2.5:7b",
                "prompt": prompt,
                "stream": False,
                "options": {
                    "temperature": 0.1,
                    "top_p": 0.9,
                    "max_tokens": 1000
                }
            },
            timeout=60
        )
        
        print(f"响应状态码: {response.status_code}")
        
        if response.status_code == 200:
            result = response.json()
            raw_response = result.get('response', '')
            
            print("\n=== 原始响应 ===")
            print(repr(raw_response))  # 使用repr显示原始字符串，包括转义字符
            
            print("\n=== 格式化响应 ===")
            print(raw_response)
            
            # 尝试解析JSON
            print("\n=== 尝试解析JSON ===")
            try:
                parsed = json.loads(raw_response)
                print("JSON解析成功!")
                print(json.dumps(parsed, ensure_ascii=False, indent=2))
            except json.JSONDecodeError as e:
                print(f"JSON解析失败: {e}")
                
                # 尝试查找JSON部分
                import re
                json_pattern = r'```json\s*(.*?)\s*```'
                json_match = re.search(json_pattern, raw_response, re.DOTALL)
                
                if json_match:
                    print("\n找到markdown中的JSON:")
                    json_str = json_match.group(1).strip()
                    print(repr(json_str))
                    try:
                        parsed = json.loads(json_str)
                        print("从markdown解析成功!")
                        print(json.dumps(parsed, ensure_ascii=False, indent=2))
                    except json.JSONDecodeError as e2:
                        print(f"从markdown解析也失败: {e2}")
                else:
                    print("未找到markdown格式的JSON")
        else:
            print(f"请求失败: {response.status_code}")
            print(response.text)
            
    except Exception as e:
        print(f"请求异常: {e}")

if __name__ == "__main__":
    debug_model_response()
