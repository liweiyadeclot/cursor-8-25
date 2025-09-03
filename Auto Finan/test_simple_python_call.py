#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简单的Python调用测试
"""

import sys
import os

def main():
    """主函数"""
    print("=== 简单Python调用测试 ===")
    print()
    
    # 检查命令行参数
    print("命令行参数:")
    for i, arg in enumerate(sys.argv):
        print(f"  {i}: {arg}")
    print()
    
    # 检查当前目录
    print(f"当前工作目录: {os.getcwd()}")
    print()
    
    # 检查文件
    files = ["config.json", "test_mouse_keyboard.py"]
    for file in files:
        if os.path.exists(file):
            print(f"✓ {file} 存在")
        else:
            print(f"✗ {file} 不存在")
    print()
    
    # 模拟成功
    print("✓ 测试成功完成")
    print("退出码: 0")
    
    return 0

if __name__ == "__main__":
    sys.exit(main())






