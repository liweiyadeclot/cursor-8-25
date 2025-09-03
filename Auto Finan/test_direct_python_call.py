#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试直接Python调用
"""

import subprocess
import sys
import os

def test_direct_python_call():
    """测试直接Python调用"""
    print("=== 测试直接Python调用 ===")
    print()
    
    # 测试命令
    cmd = [
        "python", 
        "test_mouse_keyboard.py", 
        "--config", "config.json", 
        "--folder", r"C:\Users\FH\Documents\报销单", 
        "--file", "test_file.pdf"
    ]
    
    print(f"执行命令: {' '.join(cmd)}")
    print(f"当前工作目录: {os.getcwd()}")
    print()
    
    try:
        # 执行命令
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=30
        )
        
        print(f"退出码: {result.returncode}")
        print(f"标准输出: {result.stdout}")
        print(f"错误输出: {result.stderr}")
        
        if result.returncode == 0:
            print("✓ 调用成功")
        else:
            print("✗ 调用失败")
            
    except subprocess.TimeoutExpired:
        print("✗ 执行超时")
    except Exception as e:
        print(f"✗ 执行异常: {e}")
    
    print()
    print("测试完成！")

if __name__ == "__main__":
    test_direct_python_call()






