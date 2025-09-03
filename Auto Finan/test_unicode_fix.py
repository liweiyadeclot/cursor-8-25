#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试Unicode编码修复
"""

import sys
import os
from pathlib import Path

def test_unicode_print():
    """测试Unicode字符打印"""
    print("=== 测试Unicode编码修复 ===")
    print()
    
    # 测试各种Unicode字符
    test_chars = [
        "✓",  # 原来的字符
        "❌",  # 原来的字符
        "[OK]",  # 新的字符
        "[ERROR]",  # 新的字符
        "测试中文",
        "Test English"
    ]
    
    for char in test_chars:
        try:
            print(f"测试字符: {char}")
        except UnicodeEncodeError as e:
            print(f"Unicode编码错误: {e}")
        except Exception as e:
            print(f"其他错误: {e}")
    
    print()
    print("测试完成！")

def test_folder_creation():
    """测试文件夹创建逻辑"""
    print("=== 测试文件夹创建逻辑 ===")
    print()
    
    test_folder = "test_folder_creation"
    
    try:
        folder_path_obj = Path(test_folder)
        if not folder_path_obj.exists():
            print(f"文件夹路径不存在，正在创建: {test_folder}")
            folder_path_obj.mkdir(parents=True, exist_ok=True)
            print(f"[OK] 文件夹创建成功: {test_folder}")
        else:
            print(f"[OK] 文件夹路径已存在: {test_folder}")
            
        # 清理测试文件夹
        if folder_path_obj.exists():
            folder_path_obj.rmdir()
            print(f"[OK] 测试文件夹已清理")
            
    except Exception as e:
        print(f"[ERROR] 创建文件夹失败: {e}")
    
    print()
    print("文件夹创建测试完成！")

if __name__ == "__main__":
    test_unicode_print()
    test_folder_creation()






