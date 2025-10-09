#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试直接从Excel生成MCP提示词的完整流程（跳过LLM自然语言生成）
"""

from workflow_core import process_excel_to_mcp_direct, batch_process_excel_to_mcp_direct
import json

def test_single_serial():
    """测试单个序号的直接转换"""
    excel_path = r"C:\Users\FH\source\repos\Auto Finan\Auto Finan\420财务050823.xlsx"
    sheet_name = "3-报销"
    serial = 1
    
    print("=" * 60)
    print(f"=== 测试单个序号：{serial} ===")
    print(f"Excel文件: {excel_path}")
    print(f"工作表: {sheet_name}")
    print("=" * 60)
    
    # 先获取JSON数据
    from excel_to_nl import excel_to_json_direct
    json_data = excel_to_json_direct(excel_path, sheet_name, serial)
    
    if json_data:
        print("\n=== 提取的JSON数据 ===")
        print(json.dumps(json_data, ensure_ascii=False, indent=2))
    
    # 再生成MCP提示词
    mcp_prompt = process_excel_to_mcp_direct(excel_path, sheet_name, serial)
    
    if mcp_prompt:
        print("\n=== 生成的MCP提示词 ===")
        print(mcp_prompt)
    else:
        print("\n生成失败或未找到数据")


def test_batch_processing():
    """测试批量处理所有序号"""
    excel_path = r"C:\Users\FH\source\repos\Auto Finan\Auto Finan\420财务050823.xlsx"
    sheet_name = "3-报销"
    
    print("\n" + "=" * 60)
    print(f"=== 测试批量处理 ===")
    print(f"Excel文件: {excel_path}")
    print(f"工作表: {sheet_name}")
    print("=" * 60)
    
    results = batch_process_excel_to_mcp_direct(excel_path, sheet_name)
    
    print(f"\n=== 批量处理完成 ===")
    print(f"成功处理 {len(results)} 个序号")
    
    # 显示每个序号的JSON数据和MCP提示词
    for result in results:
        print(f"\n{'='*60}")
        print(f"序号: {result['serial']}")
        print(f"\n--- JSON数据 ---")
        print(json.dumps(result['json_data'], ensure_ascii=False, indent=2))
        print(f"\n--- MCP提示词 ---")
        print(result['mcp_prompt'])


def test_travel_sheet():
    """测试差旅表的直接转换"""
    excel_path = r"C:\Users\FH\source\repos\Auto Finan\Auto Finan\420财务050823.xlsx"
    sheet_name = "3-差旅"
    serial = 1
    
    print("\n" + "=" * 60)
    print(f"=== 测试差旅表序号：{serial} ===")
    print(f"Excel文件: {excel_path}")
    print(f"工作表: {sheet_name}")
    print("=" * 60)
    
    # 先获取JSON数据
    from excel_to_nl import excel_to_json_direct
    json_data = excel_to_json_direct(excel_path, sheet_name, serial)
    
    if json_data:
        print("\n=== 提取的JSON数据 ===")
        print(json.dumps(json_data, ensure_ascii=False, indent=2))
    
    # 再生成MCP提示词
    mcp_prompt = process_excel_to_mcp_direct(excel_path, sheet_name, serial)
    
    if mcp_prompt:
        print("\n=== 生成的MCP提示词 ===")
        print(mcp_prompt)
    else:
        print("\n生成失败或未找到数据")


def test_labor_sheet():
    """测试劳务表的直接转换"""
    excel_path = r"C:\Users\FH\source\repos\Auto Finan\Auto Finan\420财务050823.xlsx"
    sheet_name = "3-劳务"
    serial = 1
    
    print("\n" + "=" * 60)
    print(f"=== 测试劳务表序号：{serial} ===")
    print(f"Excel文件: {excel_path}")
    print(f"工作表: {sheet_name}")
    print("=" * 60)
    
    # 先获取JSON数据
    from excel_to_nl import excel_to_json_direct
    json_data = excel_to_json_direct(excel_path, sheet_name, serial)
    
    if json_data:
        print("\n=== 提取的JSON数据 ===")
        print(json.dumps(json_data, ensure_ascii=False, indent=2))
    
    # 再生成MCP提示词
    mcp_prompt = process_excel_to_mcp_direct(excel_path, sheet_name, serial)
    
    if mcp_prompt:
        print("\n=== 生成的MCP提示词 ===")
        print(mcp_prompt)
    else:
        print("\n生成失败或未找到数据")


if __name__ == "__main__":
    # 测试单个序号（报销业务）
    test_single_serial()
    
    # 测试差旅表
    test_travel_sheet()
    
    # 测试劳务表
    test_labor_sheet()
    
    # 测试批量处理（可选，注释掉以避免输出过多）
    # test_batch_processing()
    
    print("\n" + "=" * 60)
    print("=== 所有测试完成 ===")
    print("=" * 60)

