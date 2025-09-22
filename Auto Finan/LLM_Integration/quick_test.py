#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速测试单个用例
"""

import json
from test_qwen_extraction import QwenExtractor

def quick_test():
    """
    快速测试单个用例
    """
    print("=== 快速测试 ===")
    
    extractor = QwenExtractor()
    
    # 检查模型状态
    if not extractor.check_model_status():
        print("错误: 模型未安装或服务未启动")
        return
    
    # 测试用例
    test_input = "张三，学工号2021001，技术部，差旅费500元，北京到上海出差，预约明天下午2点报销"
    
    print(f"输入: {test_input}")
    print("正在提取信息...")
    
    # 提取信息
    result = extractor.extract_reimbursement_info(test_input)
    
    if "error" in result:
        print(f"错误: {result['error']}")
    else:
        print("提取结果:")
        print(json.dumps(result, ensure_ascii=False, indent=2))

if __name__ == "__main__":
    quick_test()
