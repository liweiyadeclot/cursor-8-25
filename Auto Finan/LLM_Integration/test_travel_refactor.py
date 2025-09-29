#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试差旅信息重构后的功能
"""

from workflow_core import WorkflowCore

def test_travel_refactor():
    """测试差旅信息重构功能"""
    
    # 测试输入数据
    test_input = """
    业务大类：业务出差旅费
    用户名：5130008
    密码：Uestc418
    项目号：M112023ZHCG0006
    附件张数：3
    支付方式：个人转卡
    
    出差人员信息：
    1. 出差人：202422090507，姓名：陈驰，工作单位：电子科技大学，职称：无
    2. 出差人：2021090906031，姓名：高铭，工作单位：电子科技大学，职称：无
    
    费用信息：
    省份：四川除成都
    起始时间：2024-12-26
    结束时间：2024-12-27
    飞机票：0
    火车票：0
    其他交通费：36
    住宿费：0
    是否安排伙食：false
    是否安排交通：false
    
    人员信息：
    姓名：陈驰
    学工号：202422090507
    金额：100
    """
    
    print("=== 测试差旅信息重构功能 ===")
    print("输入数据：")
    print(test_input)
    print("\n" + "="*50 + "\n")
    
    # 创建WorkflowCore实例
    workflow = WorkflowCore()
    
    # 测试JSON提取
    print("1. 测试JSON提取...")
    result = workflow.extract_form_json(test_input)
    
    if "error" in result:
        print(f"JSON提取失败: {result['error']}")
        return
    
    print("提取的JSON数据：")
    import json
    print(json.dumps(result, ensure_ascii=False, indent=2))
    print("\n" + "="*50 + "\n")
    
    # 检查新的数据结构
    print("2. 检查新的数据结构...")
    
    if "travelPerson" in result:
        print("✓ travelPerson字段存在")
        travel_persons = result.get("travelPerson", [])
        print(f"  出差人员数量: {len(travel_persons)}")
        for i, person in enumerate(travel_persons):
            print(f"  人员{i+1}: {person}")
    else:
        print("✗ travelPerson字段不存在")
    
    if "travelExpenses" in result:
        print("✓ travelExpenses字段存在")
        travel_expenses = result.get("travelExpenses", [])
        print(f"  费用项目数量: {len(travel_expenses)}")
        for i, expense in enumerate(travel_expenses):
            print(f"  费用项目{i+1}: {expense}")
    else:
        print("✗ travelExpenses字段不存在")
    
    # 检查是否还有旧的travel字段
    if "travel" in result:
        print("⚠ 仍然存在旧的travel字段，可能需要清理")
    else:
        print("✓ 旧的travel字段已成功移除")
    
    print("\n" + "="*50 + "\n")
    
    # 测试字段路径收集
    print("3. 测试字段路径收集...")
    field_paths = workflow.collect_field_paths(result)
    print("收集到的字段路径：")
    for path in field_paths:
        print(f"  {path}")
    
    print("\n" + "="*50 + "\n")
    
    # 测试Playwright提示词生成
    print("4. 测试Playwright提示词生成...")
    try:
        prompt_result = workflow.build_playwright_prompt_from_input(test_input)
        if "error" in prompt_result:
            print(f"提示词生成失败: {prompt_result['error']}")
        else:
            print("生成的Playwright提示词：")
            print(prompt_result.get("prompt", ""))
    except Exception as e:
        print(f"提示词生成异常: {e}")

if __name__ == "__main__":
    test_travel_refactor()
