#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
使用workflow_core.py处理Excel数据，提取JSON并生成MCP提示词
"""

import sys
import os
from excel_to_nl import generate_single_nl_from_excel
from workflow_core import WorkflowCore

def main():
    filepath = r'C:\Users\FH\source\repos\Auto Finan\Auto Finan\420财务050823.xlsx'
    sheet_name = '3-报销'
    
    print('=== 使用workflow_core.py处理每个序号 ===')
    
    # 初始化工作流程控制器
    workflow = WorkflowCore()
    
    # 存储每个序号的处理结果
    results = []
    
    for serial in [1, 2, 3]:
        print(f'\n=== 处理序号 {serial} ===')
        
        try:
            # 1. 生成自然语言总结
            print(f'1. 生成序号 {serial} 的自然语言总结...')
            nl_text = generate_single_nl_from_excel(filepath, sheet_name, serial, use_llm=True)
            print(f'自然语言: {nl_text}')
            
            # 2. 提取JSON数据
            print(f'2. 提取序号 {serial} 的JSON数据...')
            json_data = workflow.extract_form_json(nl_text)
            print(f'提取的JSON: {json_data}')
            
            # 3. 生成MCP提示词
            print(f'3. 生成序号 {serial} 的MCP提示词...')
            mcp_prompt = workflow.build_playwright_prompt_from_data(json_data)
            print(f'MCP提示词: {mcp_prompt}')
            
            # 保存结果
            results.append({
                'serial': serial,
                'nl_text': nl_text,
                'json_data': json_data,
                'mcp_prompt': mcp_prompt
            })
            
        except Exception as e:
            print(f"处理序号 {serial} 时出错: {e}")
            import traceback
            traceback.print_exc()
            results.append({
                'serial': serial,
                'nl_text': "",
                'json_data': {},
                'mcp_prompt': "",
                'error': str(e)
            })
    
    return results

if __name__ == "__main__":
    results = main()
    
    # 打印总结
    print('\n=== 处理结果总结 ===')
    for result in results:
        print(f'\n序号 {result["serial"]}:')
        if 'error' in result:
            print(f'  错误: {result["error"]}')
        else:
            print(f'  自然语言: {result["nl_text"][:100]}...')
            print(f'  JSON数据: {result["json_data"]}')
            print(f'  MCP提示词: {result["mcp_prompt"][:200]}...')
