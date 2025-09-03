# Python鼠标键盘自动化功能使用说明

## 概述

本项目从`login_automation.py`中提取了鼠标键盘输入模拟功能，创建了独立的`mouse_keyboard_automation.py`模块，使用`pyautogui`库来实现Windows系统下的鼠标键盘自动化操作。

## 功能特性

### 1. 基本操作
- **鼠标点击**: 模拟鼠标在指定坐标的点击操作
- **键盘输入**: 模拟键盘文本输入
- **按键操作**: 模拟单个按键或组合键操作
- **鼠标移动**: 移动鼠标到指定位置
- **鼠标滚轮**: 模拟鼠标滚轮滚动

### 2. 文件保存流程
实现了完整的6步文件保存流程：
1. 根据提供的坐标，模拟鼠标点击，从而点击按钮
2. 根据提供的坐标，模拟鼠标点击，从而选择输入框
3. 根据提供的参数，在输入框中输入对应的内容，这里是文件路径
4. 根据提供的坐标，模拟鼠标点击，选择文件名输入框
5. 输入文件名
6. 根据提供的坐标，点击保存按钮

### 3. 打印对话框处理
专门针对Chrome打印对话框的处理流程，包含确认覆盖功能。

## 安装依赖

确保已安装`pyautogui`库：

```bash
pip install pyautogui
```

## 使用方法

### 1. 基本使用

```python
from mouse_keyboard_automation import create_mouse_keyboard_automation

# 创建自动化实例
automation = create_mouse_keyboard_automation()

# 鼠标点击
automation.click_mouse(100, 100, 0.5)

# 键盘输入
automation.type_text("Hello World", 0.5)

# 按键操作
automation.press_key('enter', 0.5)

# 组合键
automation.press_hotkey('ctrl', 's', delay=0.5)

# 获取鼠标位置
x, y = automation.get_mouse_position()
print(f"当前鼠标位置: ({x}, {y})")
```

### 2. 文件保存流程

```python
# 定义坐标配置
file_save_coords = {
    "button": {"x": 1562, "y": 1083},
    "filepath_input": {"x": 720, "y": 98},
    "filename_input": {"x": 404, "y": 17},
    "save_button": {"x": 971, "y": 869}
}

# 执行文件保存流程
success = automation.execute_file_save_process(
    r"C:\Users\FH\Documents\test", 
    "test_file.pdf", 
    file_save_coords
)

if success:
    print("文件保存成功")
else:
    print("文件保存失败")
```

### 3. 打印对话框处理

```python
# 使用默认配置（从config.py读取）
success = automation.execute_print_dialog_process(
    r"C:\Users\FH\Documents\pdf_output", 
    "报销单.pdf"
)

# 或者使用自定义坐标配置
custom_coords = {
    "print_button": {"x": 1562, "y": 1083},
    "filepath_input": {"x": 720, "y": 98},
    "filename_input": {"x": 404, "y": 17},
    "save_button": {"x": 971, "y": 869},
    "yes_button": {"x": 700, "y": 450}
}

success = automation.execute_print_dialog_process(
    r"C:\Users\FH\Documents\pdf_output", 
    "报销单.pdf",
    custom_coords
)
```

## 坐标获取方法

### 1. 使用Python获取坐标

```python
from mouse_keyboard_automation import create_mouse_keyboard_automation

automation = create_mouse_keyboard_automation()

# 移动鼠标到目标位置，然后获取坐标
print("请将鼠标移动到目标位置，然后按回车...")
input()
x, y = automation.get_mouse_position()
print(f"坐标: ({x}, {y})")
```

### 2. 使用Windows自带的坐标获取工具
- 按`Win + R`，输入`cmd`
- 在命令行中输入：`powershell -command "Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Cursor]::Position"`
- 移动鼠标到目标位置，按回车获取坐标

### 3. 使用第三方工具
- **AutoHotkey Window Spy**: 可以实时显示鼠标坐标
- **Screen Ruler**: 屏幕标尺工具
- **MousePos**: 专门的鼠标坐标获取工具

## 配置说明

### 坐标配置格式

```python
coordinates = {
    "print_button": {"x": 1562, "y": 1083},      # 打印按钮
    "filepath_input": {"x": 720, "y": 98},       # 文件路径输入框
    "filename_input": {"x": 404, "y": 17},       # 文件名输入框
    "save_button": {"x": 971, "y": 869},         # 保存按钮
    "yes_button": {"x": 700, "y": 450}           # 确认按钮（可选）
}
```

### 配置文件集成

模块会自动从`config.py`中读取以下配置：
- `PRINT_DIALOG_COORDINATES`: 打印对话框坐标配置
- `PRINT_FILE_PATH`: 打印文件保存路径
- `PRINT_OUTPUT_DIR`: PDF输出目录

## 安全设置

模块包含以下安全设置：
- `pyautogui.FAILSAFE = True`: 鼠标移动到屏幕左上角时停止程序
- `pyautogui.PAUSE = 0.1`: 每个操作之间的暂停时间

## 错误处理

代码包含完整的异常处理机制：
- 坐标无效时的错误提示
- 操作失败时的返回值检查
- 详细的日志输出

## 与原始代码的对比

| 功能 | login_automation.py | mouse_keyboard_automation.py |
|------|-------------------|------------------------------|
| 鼠标点击 | page.mouse.click() | pyautogui.click() |
| 键盘输入 | page.keyboard.type() | pyautogui.typewrite() |
| 按键操作 | page.keyboard.press() | pyautogui.press() |
| 组合键 | page.keyboard.press() | pyautogui.hotkey() |
| 坐标获取 | 手动配置 | pyautogui.position() |
| 错误处理 | try-except | try-except + 返回值 |
| 日志记录 | logging | logging |

## 运行示例

### 1. 直接运行演示

```bash
python mouse_keyboard_automation.py
```

### 2. 在代码中使用

```python
from mouse_keyboard_automation import create_mouse_keyboard_automation, demo_mouse_keyboard_automation

# 运行演示
demo_mouse_keyboard_automation()

# 或者创建实例使用
automation = create_mouse_keyboard_automation()
# ... 使用automation对象
```

### 3. 使用配置文件

```python
from mouse_keyboard_automation import execute_print_dialog_with_config

# 使用config.py中的配置执行
execute_print_dialog_with_config()
```

## 注意事项

1. **坐标准确性**: 坐标值必须准确，建议多次测试确认
2. **屏幕分辨率**: 不同屏幕分辨率下坐标可能不同
3. **窗口位置**: 确保目标窗口在前台且位置固定
4. **权限要求**: 某些操作可能需要管理员权限
5. **防病毒软件**: 某些防病毒软件可能会阻止模拟输入
6. **安全机制**: 鼠标移动到屏幕左上角会停止程序

## 扩展功能

可以根据需要扩展以下功能：
- 图像识别定位
- 多显示器支持
- 更复杂的操作序列
- 配置文件支持
- GUI界面

## 日志文件

模块会生成`mouse_keyboard_automation.log`日志文件，记录所有操作和错误信息。

## 故障排除

### 常见问题

1. **坐标不准确**
   - 检查屏幕分辨率
   - 确认窗口位置
   - 重新获取坐标

2. **操作失败**
   - 检查目标窗口是否在前台
   - 确认权限设置
   - 查看日志文件

3. **程序意外停止**
   - 检查是否触发了安全机制
   - 确认鼠标没有移动到屏幕左上角

### 调试技巧

```python
# 启用详细日志
import logging
logging.getLogger().setLevel(logging.DEBUG)

# 获取鼠标位置进行调试
x, y = automation.get_mouse_position()
print(f"调试：当前鼠标位置 ({x}, {y})")
```
