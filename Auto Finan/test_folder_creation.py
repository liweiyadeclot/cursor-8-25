#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试文件夹创建功能
"""

import os
from pathlib import Path

def test_folder_creation():
    """测试文件夹创建功能"""
    print("=== 测试文件夹创建功能 ===")
    
    # 测试路径
    test_paths = [
        r"C:\Users\FH\Documents\报销单",
        r"C:\Users\FH\Documents\测试文件夹\子文件夹",
        r"D:\临时测试\多层\文件夹\结构"
    ]
    
    for folder_path in test_paths:
        print(f"\n测试路径: {folder_path}")
        
        try:
            folder_path_obj = Path(folder_path)
            
            # 检查路径是否存在
            if folder_path_obj.exists():
                print(f"✓ 文件夹已存在: {folder_path}")
            else:
                print(f"文件夹不存在，正在创建: {folder_path}")
                folder_path_obj.mkdir(parents=True, exist_ok=True)
                print(f"✓ 文件夹创建成功: {folder_path}")
                
                # 验证创建是否成功
                if folder_path_obj.exists():
                    print(f"✓ 验证成功: 文件夹确实已创建")
                else:
                    print(f"❌ 验证失败: 文件夹创建后仍不存在")
                    
        except Exception as e:
            print(f"❌ 创建文件夹失败: {e}")
    
    print("\n=== 测试完成 ===")

if __name__ == "__main__":
    test_folder_creation()






