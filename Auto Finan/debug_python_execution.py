#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
调试Python脚本执行
"""

import sys
import os
import json
import time

def debug_python_execution():
    """调试Python脚本执行"""
    print("=== Python脚本执行调试 ===")
    print()
    
    # 检查Python版本
    print(f"Python版本: {sys.version}")
    print(f"Python路径: {sys.executable}")
    print()
    
    # 检查当前工作目录
    print(f"当前工作目录: {os.getcwd()}")
    print()
    
    # 检查命令行参数
    print("命令行参数:")
    for i, arg in enumerate(sys.argv):
        print(f"  {i}: {arg}")
    print()
    
    # 检查文件是否存在
    files_to_check = [
        "config.json",
        "test_mouse_keyboard.py",
        "requirements_mouse_keyboard.txt"
    ]
    
    print("文件存在性检查:")
    for file in files_to_check:
        exists = os.path.exists(file)
        print(f"  {file}: {'✓ 存在' if exists else '✗ 不存在'}")
    print()
    
    # 尝试读取配置文件
    try:
        with open('config.json', 'r', encoding='utf-8') as f:
            config_data = json.load(f)
        print("✓ 配置文件读取成功")
        print(f"配置内容: {json.dumps(config_data, indent=2, ensure_ascii=False)}")
    except Exception as e:
        print(f"✗ 配置文件读取失败: {e}")
    print()
    
    # 检查Python包
    try:
        import pyautogui
        print("✓ pyautogui 已安装")
    except ImportError:
        print("✗ pyautogui 未安装")
    
    try:
        import time
        print("✓ time 模块可用")
    except ImportError:
        print("✗ time 模块不可用")
    
    try:
        import json
        print("✓ json 模块可用")
    except ImportError:
        print("✗ json 模块不可用")
    
    try:
        import argparse
        print("✓ argparse 模块可用")
    except ImportError:
        print("✗ argparse 模块不可用")
    
    try:
        import pathlib
        print("✓ pathlib 模块可用")
    except ImportError:
        print("✗ pathlib 模块不可用")
    print()
    
    # 模拟一些操作
    print("模拟操作测试:")
    print("  1. 检查屏幕分辨率...")
    try:
        import pyautogui
        screen_width, screen_height = pyautogui.size()
        print(f"    屏幕分辨率: {screen_width}x{screen_height}")
    except Exception as e:
        print(f"    ✗ 获取屏幕分辨率失败: {e}")
    
    print("  2. 检查鼠标位置...")
    try:
        import pyautogui
        mouse_x, mouse_y = pyautogui.position()
        print(f"    鼠标位置: ({mouse_x}, {mouse_y})")
    except Exception as e:
        print(f"    ✗ 获取鼠标位置失败: {e}")
    
    print()
    print("调试完成！")

if __name__ == "__main__":
    debug_python_execution()
