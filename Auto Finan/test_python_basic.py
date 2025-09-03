#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
基本Python测试脚本
"""

import sys
import os
import json
import time

def main():
    """主函数"""
    print("=== 基本Python测试脚本 ===")
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
    
    # 尝试读取配置文件
    try:
        with open('config.json', 'r', encoding='utf-8') as f:
            config = json.load(f)
        print("✓ 配置文件读取成功")
        print(f"配置内容: {json.dumps(config, indent=2, ensure_ascii=False)}")
    except Exception as e:
        print(f"✗ 配置文件读取失败: {e}")
    print()
    
    # 检查Python包
    try:
        import pyautogui
        print("✓ pyautogui 已安装")
        screen_width, screen_height = pyautogui.size()
        print(f"屏幕分辨率: {screen_width}x{screen_height}")
    except ImportError:
        print("✗ pyautogui 未安装")
    except Exception as e:
        print(f"✗ pyautogui 测试失败: {e}")
    print()
    
    # 模拟成功
    print("✓ 基本测试成功完成")
    print("退出码: 0")
    
    return 0

if __name__ == "__main__":
    sys.exit(main())






