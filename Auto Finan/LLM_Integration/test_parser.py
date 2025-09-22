#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试JSON解析功能
"""

import json
from test_qwen_extraction import QwenExtractor

def test_parser():
    """
    测试解析功能
    """
    extractor = QwenExtractor()
    
    # 模拟模型返回的响应
    test_response = """根据您提供的用户输入信息，以下是按照要求分类整理后的JSON格式输出：

```json
{
  "login": {
    "username": null,
    "password": null
  },
  "project": {
    "projectNumber": "P2024001",
    "attachmentCount": null,
    "paymentMethod": null
  },
  "expense": {
    "category": "办公用品",
    "amount": "200元"
  },
  "personnel": {
    "studentId": null,
    "bankCard": null
  },
  "appointment": {
    "date": null,
    "location": null
  },
  "travel": {
    "name": "李四",
    "personnelType": null,
    "destination": null
  },
  "labor": {
    "laborType": null,
    "reason": null
  }
}
```

根据提供的信息，我们提取到了报销科目和部分人员信息。其他阶段没有相关描述，因此填写为null。"""
    
    print("测试响应解析...")
    print("原始响应:")
    print(test_response)
    print("\n" + "="*50 + "\n")
    
    result = extractor._parse_response(test_response)
    
    if "error" in result:
        print("解析失败:")
        print(result)
    else:
        print("解析成功:")
        print(json.dumps(result, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    test_parser()
