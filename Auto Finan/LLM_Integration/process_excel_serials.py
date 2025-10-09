#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
处理Excel中的序号，生成自然语言总结
"""

import sys
import os
from excel_to_nl import generate_single_nl_from_excel

def main():
    filepath = r'C:\Users\FH\source\repos\Auto Finan\Auto Finan\420财务050823.xlsx'
    sheet_name = '3-报销'
    
    print('=== 生成每个序号的报销信息总结 ===')
    
    # 存储每个序号的自然语言总结
    nl_summaries = []
    
    for serial in [1, 2, 3]:
        print(f'\n序号 {serial}:')
        try:
            nl_text = generate_single_nl_from_excel(filepath, sheet_name, serial, use_llm=True)
            print(nl_text)
            nl_summaries.append((serial, nl_text))
        except Exception as e:
            print(f"处理序号 {serial} 时出错: {e}")
            nl_summaries.append((serial, ""))
    
    return nl_summaries

if __name__ == "__main__":
    summaries = main()
