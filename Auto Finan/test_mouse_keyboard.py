import pyautogui
import time
import json
import argparse
import sys
import os
from pathlib import Path


class AutoClicker:
    def __init__(self, config_file):
        """
        初始化自动点击器
        :param config_file: 配置文件路径
        """
        self.config = self.load_config(config_file)
        # 设置安全暂停，将鼠标移动到角落会触发异常
        pyautogui.FAILSAFE = True
        # 每次操作后暂停0.5秒
        pyautogui.PAUSE = 0.5

    def load_config(self, config_file):
        """
        加载配置文件
        """
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            return config
        except FileNotFoundError:
            print(f"错误: 配置文件 {config_file} 不存在")
            sys.exit(1)
        except json.JSONDecodeError as e:
            print(f"错误: 配置文件格式错误 - {e}")
            sys.exit(1)

    def click_position(self, position_name, delay=1.0):
        """
        点击指定位置的坐标
        :param position_name: 配置文件中定义的位置名称
        :param delay: 点击后的延迟时间
        """
        if position_name not in self.config['ScreenPositions']:
            print(f"错误: 位置 '{position_name}' 未在配置文件中定义")
            return False

        position = self.config['ScreenPositions'][position_name]
        x, y = position['x'], position['y']

        print(f"正在点击 {position_name} - 坐标: ({x}, {y})")
        pyautogui.click(x, y)
        time.sleep(delay)
        return True

    def type_text(self, text, delay=0.1):
        """
        模拟键盘输入文本
        :param text: 要输入的文本
        :param delay: 每个字符输入后的延迟
        """
        print(f"正在输入文本: {text}")
        pyautogui.write(text, interval=delay)
        time.sleep(0.5)

    def clear_input_field(self):
        """
        清空当前输入框（Ctrl+A + Delete）
        """
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.2)
        pyautogui.press('delete')
        time.sleep(0.5)

    def execute_workflow(self, folder_path, file_name):
        """
        执行完整的工作流程
        :param folder_path: 文件夹路径
        :param file_name: 文件名
        """
        print("开始执行自动化流程...")
        print(f"文件夹路径: {folder_path}")
        print(f"文件名: {file_name}")
        print("-" * 50)

        # 检查并创建保存文件路径
        try:
            folder_path_obj = Path(folder_path)
            if not folder_path_obj.exists():
                print(f"文件夹路径不存在，正在创建: {folder_path}")
                folder_path_obj.mkdir(parents=True, exist_ok=True)
                print(f"[OK] 文件夹创建成功: {folder_path}")
            else:
                print(f"[OK] 文件夹路径已存在: {folder_path}")
        except Exception as e:
            print(f"[ERROR] 创建文件夹失败: {e}")
            return False

        try:
            # 1. 点击第一个按钮
            print("\n步骤1: 点击第一个按钮")
            if not self.click_position('first_button'):
                return False

            # 2. 点击文件路径输入框
            print("\n步骤2: 点击文件路径输入框")
            if not self.click_position('folder_input'):
                return False

            # 3. 输入文件夹路径
            print("\n步骤3: 输入文件夹路径")
            self.clear_input_field()  # 清空可能存在的文本
            self.type_text(folder_path)
            
            # 3.5. 按Enter键进入目标目录
            print("\n步骤3.5: 按Enter键进入目标目录")
            pyautogui.press('enter')
            time.sleep(1.0)  # 等待目录切换完成

            # 4. 点击文件名输入框
            print("\n步骤4: 点击文件名输入框")
            if not self.click_position('file_input'):
                return False

            # 5. 输入文件名
            print("\n步骤5: 输入文件名")
            self.clear_input_field()  # 清空可能存在的文本
            self.type_text(file_name)

            # 6. 点击确认按钮
            print("\n步骤6: 点击确认按钮")
            if not self.click_position('confirm_button'):
                return False

            print("\n[SUCCESS] 自动化流程执行完成！")
            return True

        except pyautogui.FailSafeException:
            print("\n[ERROR] 操作被用户中断（鼠标移动到屏幕角落）")
            return False
        except Exception as e:
            print(f"\n[ERROR] 执行过程中发生错误: {e}")
            return False


def main():
    # 设置命令行参数解析
    parser = argparse.ArgumentParser(description='自动化鼠标点击和键盘输入程序')
    parser.add_argument('--config', '-c', default='config.json',
                        help='配置文件路径 (默认: config.json)')
    parser.add_argument('--folder', '-f', required=True,
                        help='要输入的文件夹路径')
    parser.add_argument('--file', '-n', required=True,
                        help='要输入的文件名')
    parser.add_argument('--delay', '-d', type=float, default=1.0,
                        help='每次点击后的延迟时间 (默认: 1.0秒)')

    args = parser.parse_args()

    # 检查配置文件是否存在
    if not os.path.exists(args.config):
        print(f"错误: 配置文件 {args.config} 不存在")
        sys.exit(1)

    # 创建自动点击器实例
    clicker = AutoClicker(args.config)

    # 设置全局延迟
    pyautogui.PAUSE = args.delay

    # 执行工作流程
    print("将在5秒后开始执行，请将鼠标移动到屏幕角落可随时中断...")
    time.sleep(5)

    success = clicker.execute_workflow(args.folder, args.file)

    if success:
        print("[SUCCESS] 任务执行成功！")
        sys.exit(0)
    else:
        print("[ERROR] 任务执行失败！")
        sys.exit(1)


if __name__ == "__main__":
    main()