#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试Excel到MCP提示词的完整流程
"""

from workflow_core import process_excel_to_mcp_prompt

def main():
    """测试Excel到MCP提示词的完整流程"""
    excel_path = r"C:\Users\FH\source\repos\Auto Finan\Auto Finan\420财务050823.xlsx"
    sheet_name = "3-报销"
    
    print("=== Excel到MCP提示词完整流程测试 ===")
    print(f"Excel文件: {excel_path}")
    print(f"工作表: {sheet_name}")
    print("=" * 60)
    
    try:
        # 处理Excel文件，生成MCP提示词
        mcp_prompts = process_excel_to_mcp_prompt(excel_path, sheet_name)
        
        print(f"\n=== 处理完成 ===")
        print(f"共生成 {len(mcp_prompts)} 个MCP提示词")
        
        # 显示每个MCP提示词
        for i, prompt in enumerate(mcp_prompts):
            if prompt:
                print(f"\n--- 序号 {i+1} 的MCP提示词 ---")
                print(prompt)
            else:
                print(f"\n--- 序号 {i+1} 无有效MCP提示词 ---")
        
    except Exception as e:
        print(f"测试过程中发生错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()


