#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
鼠标键盘自动化模块
参考login_automation.py中的实现，提取出鼠标键盘输入模拟功能
"""

import asyncio
import logging
import os
import time
import argparse
import json
import sys
from typing import Optional, Dict, Any, Tuple
import pyautogui
from config import PRINT_DIALOG_COORDINATES, PRINT_FILE_PATH, PRINT_OUTPUT_DIR

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('mouse_keyboard_automation.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class MouseKeyboardAutomation:
    """鼠标键盘自动化类"""
    
    def __init__(self):
        """初始化鼠标键盘自动化"""
        # 设置pyautogui的安全设置
        pyautogui.FAILSAFE = True  # 鼠标移动到屏幕左上角时停止
        pyautogui.PAUSE = 0.1      # 每个操作之间的暂停时间
        
        logger.info("鼠标键盘自动化模块初始化完成")
    
    def click_mouse(self, x: int, y: int, delay: float = 0.5) -> bool:
        """
        模拟鼠标点击指定坐标
        
        Args:
            x: X坐标
            y: Y坐标
            delay: 点击后等待时间（秒）
            
        Returns:
            bool: 是否成功点击
        """
        try:
            logger.info(f"点击鼠标坐标: ({x}, {y})")
            pyautogui.click(x, y)
            time.sleep(delay)
            logger.info(f"✓ 成功点击坐标 ({x}, {y})")
            return True
        except Exception as e:
            logger.error(f"点击鼠标失败: {e}")
            return False
    
    def type_text(self, text: str, delay: float = 0.5) -> bool:
        """
        模拟键盘输入文本
        
        Args:
            text: 要输入的文本
            delay: 输入后等待时间（秒）
            
        Returns:
            bool: 是否成功输入
        """
        try:
            logger.info(f"输入文本: {text}")
            pyautogui.typewrite(text)
            time.sleep(delay)
            logger.info(f"✓ 成功输入文本: {text}")
            return True
        except Exception as e:
            logger.error(f"输入文本失败: {e}")
            return False
    
    def press_key(self, key: str, delay: float = 0.5) -> bool:
        """
        模拟键盘按键
        
        Args:
            key: 要按的键
            delay: 按键后等待时间（秒）
            
        Returns:
            bool: 是否成功按键
        """
        try:
            logger.info(f"按下按键: {key}")
            pyautogui.press(key)
            time.sleep(delay)
            logger.info(f"✓ 成功按下按键: {key}")
            return True
        except Exception as e:
            logger.error(f"按键失败: {e}")
            return False
    
    def press_hotkey(self, *keys, delay: float = 0.5) -> bool:
        """
        模拟组合键
        
        Args:
            *keys: 要按的组合键
            delay: 按键后等待时间（秒）
            
        Returns:
            bool: 是否成功按键
        """
        try:
            logger.info(f"按下组合键: {' + '.join(keys)}")
            pyautogui.hotkey(*keys)
            time.sleep(delay)
            logger.info(f"✓ 成功按下组合键: {' + '.join(keys)}")
            return True
        except Exception as e:
            logger.error(f"组合键失败: {e}")
            return False
    
    def clear_input_field(self) -> bool:
        """
        清空输入框内容
        
        Returns:
            bool: 是否成功清空
        """
        try:
            logger.info("清空输入框内容")
            # 全选现有内容
            self.press_hotkey('ctrl', 'a')
            time.sleep(0.2)
            # 删除现有内容
            self.press_key('delete')
            time.sleep(0.2)
            logger.info("✓ 成功清空输入框内容")
            return True
        except Exception as e:
            logger.error(f"清空输入框失败: {e}")
            return False
    
    def execute_file_save_process(self, file_path: str, file_name: str, 
                                 coordinates: Dict[str, Dict[str, int]]) -> bool:
        """
        执行文件保存流程（6步流程）
        
        Args:
            file_path: 文件路径
            file_name: 文件名
            coordinates: 坐标配置字典
            
        Returns:
            bool: 是否成功执行
        """
        try:
            logger.info("开始执行文件保存流程...")
            
            # 步骤1: 根据提供的坐标，模拟鼠标点击，从而点击按钮
            logger.info("步骤1: 点击按钮")
            if not self.click_mouse(coordinates["button"]["x"], coordinates["button"]["y"], 1.0):
                return False
            
            # 步骤2: 根据提供的坐标，模拟鼠标点击，从而选择输入框
            logger.info("步骤2: 选择文件路径输入框")
            if not self.click_mouse(coordinates["filepath_input"]["x"], coordinates["filepath_input"]["y"], 0.5):
                return False
            
            # 步骤3: 根据提供的参数，在输入框中输入对应的内容，这里是文件路径
            logger.info("步骤3: 输入文件路径")
            if not self.clear_input_field():
                return False
            if not self.type_text(file_path, 0.5):
                return False
            
            # 步骤4: 根据提供的坐标，模拟鼠标点击，选择文件名输入框
            logger.info("步骤4: 选择文件名输入框")
            if not self.click_mouse(coordinates["filename_input"]["x"], coordinates["filename_input"]["y"], 0.5):
                return False
            
            # 步骤5: 输入文件名
            logger.info("步骤5: 输入文件名")
            if not self.clear_input_field():
                return False
            if not self.type_text(file_name, 0.5):
                return False
            
            # 步骤6: 根据提供的坐标，点击保存按钮
            logger.info("步骤6: 点击保存按钮")
            if not self.click_mouse(coordinates["save_button"]["x"], coordinates["save_button"]["y"], 1.0):
                return False
            
            logger.info("✓ 文件保存流程执行完成！")
            return True
            
        except Exception as e:
            logger.error(f"文件保存流程执行失败: {e}")
            return False
    
    def execute_print_dialog_process(self, file_path: str, file_name: str, 
                                   coordinates: Optional[Dict[str, Dict[str, int]]] = None) -> bool:
        """
        执行打印对话框处理流程
        
        Args:
            file_path: 文件保存路径
            file_name: 文件名
            coordinates: 坐标配置字典，如果为None则使用config.py中的配置
            
        Returns:
            bool: 是否成功执行
        """
        try:
            logger.info("开始执行打印对话框处理流程...")
            
            # 使用默认坐标配置或传入的坐标配置
            if coordinates is None:
                coordinates = PRINT_DIALOG_COORDINATES
            
            # 等待打印对话框出现
            logger.info("等待打印对话框加载...")
            time.sleep(3)
            
            # 步骤1: 点击打印按钮
            logger.info("步骤1: 点击打印按钮")
            if not self.click_mouse(coordinates["print_button"]["x"], coordinates["print_button"]["y"], 1.0):
                return False
            
            # 步骤2: 选择文件路径输入框
            logger.info("步骤2: 选择文件路径输入框")
            if not self.click_mouse(coordinates["filepath_input"]["x"], coordinates["filepath_input"]["y"], 0.5):
                return False
            
            # 步骤3: 输入文件路径
            logger.info("步骤3: 输入文件路径")
            if not self.clear_input_field():
                return False
            if not self.type_text(file_path, 0.5):
                return False
            
            # 步骤4: 选择文件名输入框
            logger.info("步骤4: 选择文件名输入框")
            if not self.click_mouse(coordinates["filename_input"]["x"], coordinates["filename_input"]["y"], 0.5):
                return False
            
            # 步骤5: 输入文件名
            logger.info("步骤5: 输入文件名")
            if not self.clear_input_field():
                return False
            if not self.type_text(file_name, 0.5):
                return False
            
            # 步骤6: 点击保存按钮
            logger.info("步骤6: 点击保存按钮")
            if not self.click_mouse(coordinates["save_button"]["x"], coordinates["save_button"]["y"], 1.0):
                return False
            
            # 步骤7: 如果有"是"按钮坐标，点击确认覆盖
            if "yes_button" in coordinates and coordinates["yes_button"]["x"] > 0:
                logger.info("步骤7: 点击确认覆盖按钮")
                time.sleep(1)  # 等待确认对话框出现
                if not self.click_mouse(coordinates["yes_button"]["x"], coordinates["yes_button"]["y"], 0.5):
                    logger.warning("确认覆盖按钮点击失败，但继续执行")
            
            logger.info("✓ 打印对话框处理流程执行完成！")
            return True
            
        except Exception as e:
            logger.error(f"打印对话框处理流程执行失败: {e}")
            return False
    
    def get_mouse_position(self) -> Tuple[int, int]:
        """
        获取当前鼠标位置
        
        Returns:
            Tuple[int, int]: (x, y)坐标
        """
        try:
            x, y = pyautogui.position()
            logger.info(f"当前鼠标位置: ({x}, {y})")
            return x, y
        except Exception as e:
            logger.error(f"获取鼠标位置失败: {e}")
            return 0, 0
    
    def move_mouse_to(self, x: int, y: int, duration: float = 0.5) -> bool:
        """
        移动鼠标到指定位置
        
        Args:
            x: 目标X坐标
            y: 目标Y坐标
            duration: 移动持续时间（秒）
            
        Returns:
            bool: 是否成功移动
        """
        try:
            logger.info(f"移动鼠标到坐标: ({x}, {y})")
            pyautogui.moveTo(x, y, duration=duration)
            logger.info(f"✓ 成功移动鼠标到坐标 ({x}, {y})")
            return True
        except Exception as e:
            logger.error(f"移动鼠标失败: {e}")
            return False
    
    def scroll(self, clicks: int, x: Optional[int] = None, y: Optional[int] = None) -> bool:
        """
        模拟鼠标滚轮滚动
        
        Args:
            clicks: 滚动次数（正数向上，负数向下）
            x: 滚动位置的X坐标（可选）
            y: 滚动位置的Y坐标（可选）
            
        Returns:
            bool: 是否成功滚动
        """
        try:
            if x is not None and y is not None:
                logger.info(f"在坐标 ({x}, {y}) 滚动鼠标滚轮: {clicks} 次")
                pyautogui.scroll(clicks, x=x, y=y)
            else:
                logger.info(f"滚动鼠标滚轮: {clicks} 次")
                pyautogui.scroll(clicks)
            
            time.sleep(0.5)
            logger.info(f"✓ 成功滚动鼠标滚轮: {clicks} 次")
            return True
        except Exception as e:
            logger.error(f"滚动鼠标滚轮失败: {e}")
            return False

def create_mouse_keyboard_automation() -> MouseKeyboardAutomation:
    """
    创建鼠标键盘自动化实例
    
    Returns:
        MouseKeyboardAutomation: 鼠标键盘自动化实例
    """
    return MouseKeyboardAutomation()

def check_environment():
    """检查Python环境"""
    try:
        import pyautogui
        print("Python环境正常")
        return True
    except ImportError as e:
        print(f"Python环境检查失败: {e}")
        return False

def get_mouse_position():
    """获取鼠标位置"""
    try:
        automation = create_mouse_keyboard_automation()
        x, y = automation.get_mouse_position()
        print(f"({x}, {y})")
        return True
    except Exception as e:
        print(f"获取鼠标位置失败: {e}")
        return False

def execute_file_save(file_path: str, file_name: str, coordinates_str: str):
    """执行文件保存流程"""
    try:
        automation = create_mouse_keyboard_automation()
        
        # 解析坐标配置
        coordinates = json.loads(coordinates_str)
        
        success = automation.execute_file_save_process(file_path, file_name, coordinates)
        if success:
            print("✓ 文件保存流程执行成功")
        else:
            print("✗ 文件保存流程执行失败")
        return success
    except Exception as e:
        print(f"执行文件保存流程失败: {e}")
        return False

def execute_print_dialog(file_path: str, file_name: str):
    """执行打印对话框处理"""
    try:
        automation = create_mouse_keyboard_automation()
        
        success = automation.execute_print_dialog_process(file_path, file_name)
        if success:
            print("✓ 打印对话框处理执行成功")
        else:
            print("✗ 打印对话框处理执行失败")
        return success
    except Exception as e:
        print(f"执行打印对话框处理失败: {e}")
        return False

def demo_mouse_keyboard_automation():
    """
    演示鼠标键盘自动化功能
    """
    print("=== 鼠标键盘自动化功能演示 ===")
    
    automation = create_mouse_keyboard_automation()
    
    # 示例1: 基本的鼠标点击和键盘输入
    print("\n--- 示例1: 基本操作 ---")
    automation.click_mouse(100, 100, 0.5)
    automation.type_text("Hello World", 0.5)
    automation.press_key('enter', 0.5)
    
    # 示例2: 文件保存流程
    print("\n--- 示例2: 文件保存流程 ---")
    file_save_coords = {
        "button": {"x": 1562, "y": 1083},
        "filepath_input": {"x": 720, "y": 98},
        "filename_input": {"x": 404, "y": 17},
        "save_button": {"x": 971, "y": 869}
    }
    
    success = automation.execute_file_save_process(
        r"C:\Users\FH\Documents\test", 
        "test_file.pdf", 
        file_save_coords
    )
    
    if success:
        print("✓ 文件保存流程演示成功")
    else:
        print("✗ 文件保存流程演示失败")
    
    # 示例3: 打印对话框处理流程
    print("\n--- 示例3: 打印对话框处理流程 ---")
    success = automation.execute_print_dialog_process(
        r"C:\Users\FH\Documents\pdf_output", 
        "报销单.pdf"
    )
    
    if success:
        print("✓ 打印对话框处理流程演示成功")
    else:
        print("✗ 打印对话框处理流程演示失败")
    
    print("演示完成！")

def execute_print_dialog_with_config():
    """
    使用配置文件中的坐标执行打印对话框处理
    """
    print("=== 使用配置坐标执行打印对话框处理 ===")
    
    automation = create_mouse_keyboard_automation()
    
    # 使用配置文件中的路径
    file_path = PRINT_FILE_PATH
    file_name = f"报销单_{time.strftime('%Y%m%d_%H%M%S')}.pdf"
    
    try:
        success = automation.execute_print_dialog_process(file_path, file_name)
        if success:
            print(f"✓ 文件已保存到: {os.path.join(file_path, file_name)}")
        else:
            print("✗ 执行失败")
    except Exception as e:
        print(f"执行失败: {e}")

def execute_print_confirm_button_click(x: int, y: int):
    """
    执行打印确认单按钮点击
    
    Args:
        x: X坐标
        y: Y坐标
    """
    print(f"=== 执行打印确认单按钮点击 ===")
    print(f"目标坐标: ({x}, {y})")
    
    automation = create_mouse_keyboard_automation()
    
    try:
        # 点击打印确认单按钮
        success = automation.click_mouse(x, y, delay=1.0)
        
        if success:
            print(f"✓ 成功点击打印确认单按钮坐标 ({x}, {y})")
            
            # 等待页面响应
            print("等待页面响应...")
            time.sleep(2)
            
            # 可以在这里添加更多的自动化操作，比如处理弹出的对话框等
            print("✓ 打印确认单按钮点击完成")
            return True
        else:
            print(f"✗ 点击打印确认单按钮失败")
            return False
            
    except Exception as e:
        print(f"执行打印确认单按钮点击时出错: {e}")
        return False

def execute_print_dialog_automation():
    """
    执行打印对话框自动化处理
    这个方法会在打印确认单按钮点击后调用，处理可能弹出的打印对话框
    """
    print(f"=== 执行打印对话框自动化处理 ===")
    
    automation = create_mouse_keyboard_automation()
    
    try:
        # 等待打印对话框出现
        print("等待打印对话框加载...")
        time.sleep(3)
        
        # 这里可以添加处理打印对话框的逻辑
        # 例如：点击保存按钮、选择文件路径、输入文件名等
        
        # 示例：处理打印对话框的保存操作
        print("开始处理打印对话框...")
        
        # 可以在这里添加具体的坐标和操作
        # 这些坐标需要根据实际的打印对话框界面来调整
        
        # 示例操作（需要根据实际情况调整坐标）
        # automation.click_mouse(800, 600, 1.0)  # 点击保存按钮
        # automation.click_mouse(700, 500, 0.5)  # 选择文件路径输入框
        # automation.type_text(r"C:\Users\FH\Documents\报销单.pdf", 0.5)  # 输入文件路径
        # automation.press_key('enter', 1.0)  # 按回车确认
        
        print("✓ 打印对话框自动化处理完成")
        return True
        
    except Exception as e:
        print(f"执行打印对话框自动化处理时出错: {e}")
        return False

def main():
    """主函数 - 处理命令行参数"""
    parser = argparse.ArgumentParser(description='鼠标键盘自动化工具')
    parser.add_argument('--operation', type=str, help='操作类型')
    parser.add_argument('--filepath', type=str, help='文件路径')
    parser.add_argument('--filename', type=str, help='文件名')
    parser.add_argument('--coordinates', type=str, help='坐标配置（JSON格式）')
    parser.add_argument('--x', type=int, help='X坐标')
    parser.add_argument('--y', type=int, help='Y坐标')
    parser.add_argument('--check', action='store_true', help='检查Python环境')
    parser.add_argument('--demo', action='store_true', help='运行演示')
    parser.add_argument('--config', action='store_true', help='使用配置文件执行')
    
    args = parser.parse_args()
    
    # 检查Python环境
    if args.check:
        check_environment()
        return
    
    # 运行演示
    if args.demo:
        demo_mouse_keyboard_automation()
        return
    
    # 使用配置文件执行
    if args.config:
        execute_print_dialog_with_config()
        return
    
    # 处理坐标参数（用于打印确认单按钮点击）
    if args.x is not None and args.y is not None:
        execute_print_confirm_button_click(args.x, args.y)
        return
    
    # 处理打印对话框自动化（无参数调用）
    if not args.operation and not args.check and not args.demo and not args.config:
        # 如果没有指定任何参数，默认执行打印对话框自动化
        execute_print_dialog_automation()
        return
    
    # 根据操作类型执行相应功能
    if args.operation:
        if args.operation == "get_position":
            get_mouse_position()
        elif args.operation == "file_save":
            if args.filepath and args.filename and args.coordinates:
                execute_file_save(args.filepath, args.filename, args.coordinates)
            else:
                print("错误：文件保存操作需要提供filepath、filename和coordinates参数")
        elif args.operation == "print_dialog":
            if args.filepath and args.filename:
                execute_print_dialog(args.filepath, args.filename)
            else:
                print("错误：打印对话框操作需要提供filepath和filename参数")
        else:
            print(f"未知操作类型: {args.operation}")
    else:
        # 默认运行演示
        demo_mouse_keyboard_automation()

if __name__ == "__main__":
    main()
