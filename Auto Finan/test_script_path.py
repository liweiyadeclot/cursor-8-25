#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试脚本路径查找
"""

import sys
import os

def main():
    """主函数"""
    print("=== 测试脚本路径查找 ===")
    print()
    
    # 检查当前目录
    print(f"当前工作目录: {os.getcwd()}")
    print()
    
    # 检查文件
    files = ["test_mouse_keyboard.py", "config.json"]
    for file in files:
        if os.path.exists(file):
            print(f"✓ {file} 存在")
        else:
            print(f"✗ {file} 不存在")
    print()
    
    # 测试Unicode字符
    try:
        print("测试Unicode字符输出...")
        print("[OK] 测试成功")
        print("[ERROR] 测试错误")
        print("✓ 测试成功")
        print("❌ 测试失败")
        print("✅ 测试完成")
        print("🎉 测试通过")
    except UnicodeEncodeError as e:
        print(f"Unicode编码错误: {e}")
    except Exception as e:
        print(f"其他错误: {e}")
    
    print()
    print("测试完成！")
    return 0

if __name__ == "__main__":
    sys.exit(main())






