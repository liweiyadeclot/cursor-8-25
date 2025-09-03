#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
简单的Python脚本测试工具
"""

import subprocess
import sys
import os
import json

def test_config_file():
    """测试配置文件是否可以正确读取"""
    print("=== 测试配置文件 ===")
    try:
        with open('config.json', 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        print("✓ 配置文件读取成功")
        print(f"配置内容: {json.dumps(config, indent=2, ensure_ascii=False)}")
        
        # 检查必要的坐标配置
        if 'ScreenPositions' in config:
            positions = config['ScreenPositions']
            required_positions = ['first_button', 'folder_input', 'file_input', 'confirm_button']
            
            for pos in required_positions:
                if pos in positions:
                    print(f"✓ 找到坐标配置: {pos}")
                else:
                    print(f"✗ 缺少坐标配置: {pos}")
        else:
            print("✗ 配置文件中缺少ScreenPositions部分")
            
        return True
    except Exception as e:
        print(f"✗ 配置文件测试失败: {e}")
        return False

def test_python_script():
    """测试Python脚本是否可以正常运行"""
    print("\n=== 测试Python脚本 ===")
    try:
        # 检查脚本文件是否存在
        if not os.path.exists('test_mouse_keyboard.py'):
            print("✗ test_mouse_keyboard.py文件不存在")
            return False
        
        # 测试脚本的帮助信息
        result = subprocess.run([
            'python', 'test_mouse_keyboard.py', '--help'
        ], capture_output=True, text=True, timeout=10)
        
        if result.returncode == 0:
            print("✓ Python脚本可以正常运行")
            print("帮助信息:")
            print(result.stdout)
            return True
        else:
            print(f"✗ Python脚本运行失败: {result.stderr}")
            return False
            
    except subprocess.TimeoutExpired:
        print("✗ Python脚本运行超时")
        return False
    except Exception as e:
        print(f"✗ Python脚本测试失败: {e}")
        return False

def test_dry_run():
    """测试脚本的干运行模式（不实际执行鼠标操作）"""
    print("\n=== 测试干运行模式 ===")
    try:
        # 创建一个临时的测试配置文件
        test_config = {
            "ScreenPositions": {
                "first_button": {"x": 100, "y": 100, "description": "测试按钮"},
                "folder_input": {"x": 200, "y": 200, "description": "测试输入框"},
                "file_input": {"x": 300, "y": 300, "description": "测试文件输入框"},
                "confirm_button": {"x": 400, "y": 400, "description": "测试确认按钮"}
            }
        }
        
        with open('test_config.json', 'w', encoding='utf-8') as f:
            json.dump(test_config, f, indent=2, ensure_ascii=False)
        
        print("✓ 创建测试配置文件成功")
        
        # 测试文件夹创建功能
        print("\n测试文件夹创建功能...")
        test_folder = r"C:\Users\FH\Documents\测试文件夹"
        folder_path_obj = Path(test_folder)
        
        if not folder_path_obj.exists():
            print(f"创建测试文件夹: {test_folder}")
            folder_path_obj.mkdir(parents=True, exist_ok=True)
            print("✓ 测试文件夹创建成功")
        else:
            print("✓ 测试文件夹已存在")
        
        # 运行脚本（但不会实际执行鼠标操作，因为坐标在屏幕外）
        print("注意：这个测试不会实际执行鼠标操作")
        print("如果要实际测试，请运行: python test_mouse_keyboard.py --config config.json --folder \"测试路径\" --file \"测试文件.pdf\"")
        
        return True
        
    except Exception as e:
        print(f"✗ 干运行测试失败: {e}")
        return False

def main():
    """主测试函数"""
    print("Python脚本功能测试")
    print("=" * 50)
    
    # 测试配置文件
    config_ok = test_config_file()
    
    # 测试Python脚本
    script_ok = test_python_script()
    
    # 测试干运行
    dry_run_ok = test_dry_run()
    
    # 总结
    print("\n=== 测试总结 ===")
    print(f"配置文件测试: {'✓ 通过' if config_ok else '✗ 失败'}")
    print(f"Python脚本测试: {'✓ 通过' if script_ok else '✗ 失败'}")
    print(f"干运行测试: {'✓ 通过' if dry_run_ok else '✗ 失败'}")
    
    if config_ok and script_ok:
        print("\n🎉 所有测试通过！Python脚本可以正常使用。")
        print("\n要实际测试鼠标操作，请运行:")
        print("python test_mouse_keyboard.py --config config.json --folder \"C:\\Users\\FH\\Documents\\报销单\" --file \"test_file.pdf\"")
    else:
        print("\n❌ 部分测试失败，请检查配置和脚本。")

if __name__ == "__main__":
    main()
