#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
标记MCP提示词为已完成的示例脚本

使用场景：
当你手动或通过Playwright MCP执行完某个MCP提示词文件后，
运行此脚本将文件名从"未预约"改为"已预约"，并同步更新Excel。
"""

from workflow_core import mark_mcp_file_as_completed

def main():
    # ===== 配置区域 =====
    EXCEL_FILE_PATH = r"C:\Users\FH\source\repos\Auto Finan\Auto Finan\420财务050823.xlsx"
    SHEET_NAME = "3-报销"  # 修改为你的工作表名
    
    # ===== 输入要标记的文件名 =====
    print("=" * 60)
    print("=== MCP提示词完成标记工具 ===")
    print("=" * 60)
    print(f"Excel文件: {EXCEL_FILE_PATH}")
    print(f"工作表: {SHEET_NAME}\n")
    
    old_filename = input("请输入要标记为已完成的文件名（如：未预约-M112023ZHCG0006-100-20251009-14-30-25.txt）：").strip()
    
    if not old_filename:
        print("错误：文件名不能为空")
        return
    
    if not old_filename.startswith("未预约-"):
        print("警告：文件名不是以'未预约-'开头，可能已经标记过了")
        confirm = input("是否继续？(y/n): ").strip().lower()
        if confirm != 'y':
            print("已取消")
            return
    
    print(f"\n正在标记文件为已完成...")
    print("-" * 60)
    
    # 调用标记函数
    new_filename = mark_mcp_file_as_completed(EXCEL_FILE_PATH, SHEET_NAME, old_filename)
    
    print("-" * 60)
    print(f"\n✅ 操作完成！")
    print(f"新文件名: {new_filename}")
    print("=" * 60)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n已取消操作")
    except Exception as e:
        print(f"\n错误: {e}")
        import traceback
        traceback.print_exc()

