#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试Enter键功能
"""

import pyautogui
import time

def test_enter_key():
    """测试Enter键功能"""
    print("=== 测试Enter键功能 ===")
    
    print("这个测试将模拟按Enter键操作")
    print("请确保当前焦点在文件路径输入框中")
    print("将在5秒后开始测试...")
    
    for i in range(5, 0, -1):
        print(f"倒计时: {i}秒")
        time.sleep(1)
    
    print("开始测试Enter键...")
    
    try:
        # 模拟按Enter键
        pyautogui.press('enter')
        print("✓ Enter键按下成功")
        
        # 等待一秒
        time.sleep(1)
        print("✓ 等待完成")
        
        print("测试完成！")
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")

if __name__ == "__main__":
    test_enter_key()






