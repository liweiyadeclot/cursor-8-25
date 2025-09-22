#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试多个科目和多个人员的处理
"""

import json
from test_qwen_extraction import QwenExtractor

def test_multiple_items():
    """
    测试多个科目和多个人员的处理
    """
    print("=== 测试多个科目和多个人员 ===")
    
    extractor = QwenExtractor()
    
    # 检查模型状态
    if not extractor.check_model_status():
        print("错误: 模型未安装或服务未启动")
        return
    
    # 测试用例：多个科目和多个人员
    test_cases = [
        {
            "name": "多个科目单个人员",
            "input": "张三，学工号2021001，差旅费300元、办公用品150元、培训费500元，项目号P2024001"
        },
        {
            "name": "多个人员单个科目",
            "input": "张三学工号2021001、李四学工号2021002、王五学工号2021003，差旅费每人500元，北京出差"
        },
        {
            "name": "多个科目多个人员",
            "input": "张三学工号2021001、李四学工号2021002，差旅费每人300元、办公用品150元、培训费每人400元，上海出差，项目号P2024002"
        }
    ]
    
    for test_case in test_cases:
        print(f"\n测试: {test_case['name']}")
        print(f"输入: {test_case['input']}")
        print("-" * 60)
        
        # 提取信息
        result = extractor.extract_reimbursement_info(test_case['input'])
        
        if "error" in result:
            print(f"错误: {result['error']}")
        else:
            print("提取结果:")
            print(json.dumps(result, ensure_ascii=False, indent=2))
            
            # 分析结果
            print("\n分析:")
            if result.get('expenses'):
                print(f"- 报销科目数量: {len(result['expenses'])}")
                for i, expense in enumerate(result['expenses']):
                    print(f"  科目{i+1}: {expense.get('category', 'N/A')} - {expense.get('amount', 'N/A')}")
            
            if result.get('personnel'):
                print(f"- 报销人员数量: {len(result['personnel'])}")
                for i, person in enumerate(result['personnel']):
                    print(f"  人员{i+1}: {person.get('name', 'N/A')} - {person.get('studentId', 'N/A')}")
            
            if result.get('travel'):
                print(f"- 出差人员数量: {len(result['travel'])}")
                for i, travel in enumerate(result['travel']):
                    print(f"  出差{i+1}: {travel.get('name', 'N/A')} - {travel.get('destination', 'N/A')}")
        
        print("\n" + "=" * 80)

if __name__ == "__main__":
    test_multiple_items()
