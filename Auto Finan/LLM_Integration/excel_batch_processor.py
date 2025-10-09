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
    import os
    import time
    
    # ===== 配置区域：可修改Excel文件路径 =====
    EXCEL_FILE_PATH = r"C:\Users\FH\source\repos\Auto Finan\Auto Finan\420财务050823.xlsx"
    OUTPUT_DIR = r"C:\Users\FH\source\repos\Auto Finan\Auto Finan\LLM_Integration\mcp_prompts"
    
    # 创建输出目录（如果不存在）
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
    
    # ===== 主程序 =====
    print("=" * 60)
    print("=== Excel批量处理器 - 生成MCP提示词 ===")
    print("=" * 60)
    print(f"Excel文件路径: {EXCEL_FILE_PATH}")
    print(f"输出目录: {OUTPUT_DIR}\n")
    
    # 提示用户输入工作表名
    sheet_name = input("请输入工作表名称（如：3-报销、3-差旅、3-劳务）：").strip()
    
    if not sheet_name:
        print("错误：工作表名不能为空")
        exit(1)
    
    print(f"\n开始处理工作表：{sheet_name}")
    print("=" * 60)
    
    try:
        # 批量处理指定工作表的所有序号
        results = batch_process_excel_to_mcp_direct(EXCEL_FILE_PATH, sheet_name)
        
        if not results:
            print(f"\n未找到任何数据或处理失败")
            exit(1)
        
        print(f"\n{'='*60}")
        print(f"=== 批量处理完成 ===")
        print(f"成功处理 {len(results)} 个序号")
        print(f"{'='*60}")
        
        # 保存每个序号的MCP提示词到文件
        saved_files = []
        for result in results:
            serial = result['serial']
            json_data = result['json_data']
            mcp_prompt = result['mcp_prompt']
            
            # 提取项目号和计算总金额
            project_number = json_data.get('project', {}).get('projectNumber', 'UNKNOWN')
            
            # 计算总金额（根据业务类型）
            total_amount = 0
            business_type = json_data.get('businessType', '')
            
            if '报销' in business_type or '差旅' in business_type:
                # 报销业务和差旅业务：从personnel数组计算
                for person in json_data.get('personnel', []):
                    amount_str = person.get('amount', '0')
                    try:
                        total_amount += float(amount_str) if amount_str else 0
                    except:
                        pass
            elif '酬金' in business_type or '劳务' in business_type:
                # 劳务业务：从laborPerson数组计算
                for person in json_data.get('laborPerson', []):
                    amount_str = person.get('singleEntryAmount', '0')
                    try:
                        total_amount += float(amount_str) if amount_str else 0
                    except:
                        pass
            
            # 生成时间戳（格式：日期-时-分-秒）
            timestamp = time.strftime("%Y%m%d-%H-%M-%S", time.localtime())
            
            # 生成文件名
            filename = f"未预约-{project_number}-{int(total_amount)}-{timestamp}.txt"
            filepath = os.path.join(OUTPUT_DIR, filename)
            
            # 保存到文件
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(mcp_prompt)
            
            saved_files.append(filepath)
            
            # 延迟1.5秒，防止同一时间覆盖
            time.sleep(1.5)
            
            # 显示摘要信息
            print(f"\n{'='*60}")
            print(f"序号: {serial}")
            print(f"项目号: {project_number}")
            print(f"总金额: {int(total_amount)}")
            print(f"已保存至: {filename}")
            print(f"MCP提示词长度: {len(mcp_prompt)} 字符")
        
        print(f"\n{'='*60}")
        print(f"=== 所有序号处理完成 ===")
        print(f"共保存 {len(saved_files)} 个MCP提示词文件")
        print(f"输出目录: {OUTPUT_DIR}")
        print(f"{'='*60}")
        
    except Exception as e:
        print(f"\n处理过程中发生错误: {e}")
        import traceback
        traceback.print_exc()
        exit(1)

