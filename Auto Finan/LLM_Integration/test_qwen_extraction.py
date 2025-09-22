#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Qwen 2.5 7B 报销信息提取测试脚本
用于测试自然语言到结构化数据的转换
"""

import requests
import json
import time
from typing import Dict, Any, Optional

class QwenExtractor:
    def __init__(self, base_url: str = "http://localhost:11434"):
        """
        初始化Qwen模型客户端
        
        Args:
            base_url: Ollama服务地址
        """
        self.base_url = base_url
        self.model_name = "qwen2.5:7b"
        
    def check_model_status(self) -> bool:
        """
        检查模型是否可用
        
        Returns:
            bool: 模型是否可用
        """
        try:
            response = requests.get(f"{self.base_url}/api/tags", timeout=5)
            if response.status_code == 200:
                models = response.json()
                model_list = [model['name'] for model in models.get('models', [])]
                return self.model_name in model_list
            return False
        except Exception as e:
            print(f"检查模型状态失败: {e}")
            return False
    
    def extract_reimbursement_info(self, user_input: str) -> Dict[str, Any]:
        """
        从用户输入中提取报销信息
        
        Args:
            user_input: 用户自然语言输入
            
        Returns:
            Dict: 提取的结构化信息
        """
        prompt = self._build_extraction_prompt(user_input)
        
        try:
            response = requests.post(
                f"{self.base_url}/api/generate",
                json={
                    "model": self.model_name,
                    "prompt": prompt,
                    "stream": False,
                    "options": {
                        "temperature": 0.1,
                        "top_p": 0.9,
                        "max_tokens": 1000
                    }
                },
                timeout=30
            )
            
            if response.status_code == 200:
                result = response.json()
                return self._parse_response(result.get('response', ''))
            else:
                return {"error": f"请求失败: {response.status_code}"}
                
        except Exception as e:
            return {"error": f"提取失败: {e}"}
    
    def _build_extraction_prompt(self, user_input: str) -> str:
        """
        构建信息提取提示词
        
        Args:
            user_input: 用户输入
            
        Returns:
            str: 完整的提示词
        """
        prompt = """你是一个专业的财务报销信息提取助手。请从用户输入的自然语言中提取报销相关信息，并按照以下7个阶段进行分类整理。

提取阶段说明：
1. 登录阶段：用户名、密码
2. 项目信息：项目号、附件张数、支付方式
3. 报销科目：科目类型、金额（可能有多个科目）
4. 报销人员：学工号、银行卡信息、个人报销金额（可能有多个人员）
5. 预约报销：报销时间、地点
6. 差旅信息：姓名、人员类型、出差地点（可能有多个出差人员）
7. 劳务信息：劳务费类型、发放事由（可能有多个劳务项目）

业务大类说明：
请根据报销内容判断业务大类，可选业务大类包括：
- "报销业务"：包含出差、交通、住宿、餐饮等费用
- "差旅业务"：包含办公设备、文具、耗材等费用
- "劳务业务"：包含培训课程、会议、学习材料等费用
业务大类默认值为报销业务

重要说明：
- 报销科目、报销人员、差旅信息、劳务信息可能是多个，请以数组形式输出
- 每个科目/人员/出差/劳务项目都是一个独立的对象
- 如果某个阶段没有相关信息，请在该阶段中填写null或空数组[]

重要：请严格按照以下JSON格式输出，使用英文键名，不要包含任何其他文字说明。

{
  "businessType": "业务大类",
  "login": {
    "username": "用户名",
    "password": "密码"
  },
  "project": {
    "projectNumber": "项目号",
    "attachmentCount": "附件张数",
    "paymentMethod": "支付方式"
  },
  "expenses": [
    {
      "category": "科目类型",
      "amount": "金额"
    }
  ],
  "personnel": [
    {
      "name": "姓名",
      "studentId": "学工号",
      "bankCard": "银行卡信息",
      "amount": "个人金额"
    }
  ],
  "appointment": {
    "date": "报销时间",
    "location": "地点"
  },
  "travel": [
    {
      "name": "姓名",
      "personnelType": "人员类型",
      "destination": "出差地点"

    }
  ],
  "labor": [
    {
      "laborType": "劳务费类型",
      "amount": "金额",
      "reason": "发放事由"
    }
  ]
}

用户输入：""" + user_input + """

JSON输出："""
        
        return prompt
    
    def _parse_response(self, response: str) -> Dict[str, Any]:
        """
        解析模型响应
        
        Args:
            response: 模型原始响应
            
        Returns:
            Dict: 解析后的结构化数据
        """
        try:
            # 尝试直接解析JSON
            if response.strip().startswith('{'):
                return json.loads(response)
            
            # 查找markdown代码块中的JSON
            import re
            
            # 匹配 ```json 和 ``` 之间的内容
            json_pattern = r'```json\s*(.*?)\s*```'
            json_match = re.search(json_pattern, response, re.DOTALL)
            
            if json_match:
                json_str = json_match.group(1).strip()
                return json.loads(json_str)
            
            # 匹配 ``` 和 ``` 之间的内容（没有json标记）
            code_pattern = r'```\s*(.*?)\s*```'
            code_match = re.search(code_pattern, response, re.DOTALL)
            
            if code_match:
                code_str = code_match.group(1).strip()
                if code_str.startswith('{'):
                    return json.loads(code_str)
            
            # 查找独立的JSON对象
            json_object_pattern = r'\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\}'
            json_matches = re.findall(json_object_pattern, response, re.DOTALL)
            
            for match in json_matches:
                try:
                    return json.loads(match)
                except json.JSONDecodeError:
                    continue
            
            # 如果响应包含其他文本，尝试提取JSON部分
            lines = response.strip().split('\n')
            for line in lines:
                line = line.strip()
                if line.startswith('{') and line.endswith('}'):
                    return json.loads(line)
            
            # 如果无法解析，返回原始响应
            return {"raw_response": response, "parse_error": "无法解析JSON格式"}
            
        except json.JSONDecodeError as e:
            return {"raw_response": response, "parse_error": f"JSON解析错误: {e}"}
        except Exception as e:
            return {"raw_response": response, "parse_error": f"解析错误: {e}"}

def test_extraction():
    """
    测试信息提取功能
    """
    print("=== Qwen 2.5 7B 报销信息提取测试 ===\n")
    
    # 创建提取器实例
    extractor = QwenExtractor()
    
    # 检查模型状态
    print("检查模型状态...")
    if not extractor.check_model_status():
        print(f"错误: 模型 {extractor.model_name} 未安装或服务未启动")
        print("请运行以下命令:")
        print("1. ollama pull qwen2.5:7b")
        print("2. ollama serve")
        return
    
    print("✓ 模型状态正常\n")
    
    # 测试用例
    test_cases = [
        "使用陈驰账户进行报销，业务大类为报销业务，账号202422090507，密码12345，报销项目号M112023ZHCG0006，附件张数为3张，支付方式为个人转卡。报销科目为办公费，金额500元。报销人1姓名张三，学工号2021001，银行卡号123，张三报销100元。报销人2姓名陈驰，学工号202422090507，银行卡号561。报销金额400元。报销时间2024年1月15日，地点清水河校区"
    ]
    
    for i, test_input in enumerate(test_cases, 1):
        print(f"测试用例 {i}:")
        print(f"输入: {test_input}")
        print("-" * 50)
        
        # 提取信息
        result = extractor.extract_reimbursement_info(test_input)
        
        # 显示结果
        if "error" in result:
            print(f"错误: {result['error']}")
        else:
            print("提取结果:")
            print(json.dumps(result, ensure_ascii=False, indent=2))
        
        print("\n" + "=" * 80 + "\n")
        
        # 添加延迟避免请求过快
        time.sleep(1)

def interactive_test():
    """
    交互式测试
    """
    print("=== 交互式测试模式 ===")
    print("输入 'quit' 退出测试\n")
    
    extractor = QwenExtractor()
    
    if not extractor.check_model_status():
        print(f"错误: 模型 {extractor.model_name} 未安装或服务未启动")
        return
    
    while True:
        user_input = input("请输入报销信息: ").strip()
        
        if user_input.lower() == 'quit':
            break
        
        if not user_input:
            continue
        
        print("正在提取信息...")
        result = extractor.extract_reimbursement_info(user_input)
        
        if "error" in result:
            print(f"错误: {result['error']}")
        else:
            print("提取结果:")
            print(json.dumps(result, ensure_ascii=False, indent=2))
        
        print("\n" + "-" * 50 + "\n")

if __name__ == "__main__":
    try:
        # 运行自动测试
        test_extraction()
        
        # 询问是否进行交互式测试
        print("是否进行交互式测试？(y/n): ", end="")
        if input().lower().startswith('y'):
            interactive_test()
            
    except KeyboardInterrupt:
        print("\n测试已取消")
    except Exception as e:
        print(f"测试过程中发生错误: {e}")
