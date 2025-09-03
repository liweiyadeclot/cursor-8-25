# 打印确认单按钮集成功能

## 功能概述

这个功能实现了在点击"打印确认单"按钮后，自动调用Python脚本处理后续的鼠标键盘操作（如处理打印对话框）。

## 工作流程

1. **通过Playwright点击按钮**：程序通过Playwright在网页上点击"打印确认单"按钮
2. **等待页面响应**：等待3秒让页面响应和可能的对话框加载
3. **调用Python脚本**：自动调用`test_mouse_keyboard.py`脚本处理后续操作
4. **处理打印对话框**：Python脚本使用`config.json`中的坐标配置执行鼠标键盘操作

## 使用方法

### 1. 在Excel中配置

在Excel文件的"打印确认单"列中设置值为`$点击`：

```
| 打印确认单 |
|-----------|
| $点击     |
```

### 2. 程序执行流程

当程序执行到这一行时：

1. 检测到"打印确认单"列的值为"$点击"
2. 通过Playwright在网页上查找并点击打印确认单按钮
3. 等待页面响应（3秒）
4. 调用Python脚本`test_mouse_keyboard.py`
5. Python脚本使用`config.json`中的坐标配置执行鼠标键盘操作

### 3. Python脚本处理

Python脚本会：

1. 等待打印对话框或其他界面出现
2. 根据预设的坐标执行鼠标点击和键盘输入操作
3. 处理文件保存、路径选择等操作

## 配置说明

### 修改打印对话框处理逻辑

如果需要修改Python脚本中的打印对话框处理逻辑，请编辑`test_mouse_keyboard.py`文件中的`execute_workflow()`方法：

```python
def execute_workflow(self, folder_path, file_name):
    """
    执行完整的工作流程
    """
    print("开始执行自动化流程...")
    
    try:
        # 1. 点击第一个按钮
        if not self.click_position('first_button'):
            return False

        # 2. 点击文件路径输入框
        if not self.click_position('folder_input'):
            return False

        # 3. 输入文件夹路径
        self.clear_input_field()
        self.type_text(folder_path)

        # 4. 点击文件名输入框
        if not self.click_position('file_input'):
            return False

        # 5. 输入文件名
        self.clear_input_field()
        self.type_text(file_name)

        # 6. 点击确认按钮
        if not self.click_position('confirm_button'):
            return False

        return True
    except Exception as e:
        print(f"执行失败: {e}")
        return False
```

### 调整等待时间

如果需要调整等待时间，可以修改：

1. **Program.cs中的等待时间**：修改`HandlePrintConfirmButton()`方法中的`await Task.Delay(3000)`
2. **Python脚本中的等待时间**：修改`execute_workflow()`方法中的延迟时间

## 错误处理

### 常见问题

1. **Python脚本执行器初始化失败**
   - 确保系统已安装Python
   - 检查Python是否在PATH环境变量中

2. **打印对话框处理失败**
   - 检查config.json中的坐标配置是否正确
   - 确认打印对话框是否正常弹出
   - 查看Python脚本的日志输出

3. **按钮点击失败**
   - 检查网页元素是否正确加载
   - 确认按钮ID或选择器是否正确

### 调试方法

1. **查看控制台输出**：程序会输出详细的执行日志
2. **检查Python日志**：查看`mouse_keyboard_automation.log`文件
3. **手动测试**：先手动点击按钮，确认打印对话框正常弹出

## 测试

运行测试脚本：

```bash
test_print_confirm_workflow.bat
```

这个脚本会：
1. 编译项目
2. 测试Python脚本执行器
3. 验证打印对话框自动化功能

## 扩展功能

### 添加更多自动化操作

可以在`execute_print_dialog_automation()`方法中添加更多的自动化操作：

```python
# 示例：处理文件覆盖确认对话框
automation.click_mouse(900, 700, 1.0)  # 点击"是"按钮

# 示例：处理其他类型的对话框
automation.press_key('enter', 0.5)  # 按回车键
```

### 支持不同的打印对话框

可以根据不同的打印对话框类型，添加不同的处理逻辑：

```python
def execute_print_dialog_automation(dialog_type="default"):
    if dialog_type == "chrome":
        # Chrome浏览器的打印对话框处理
        pass
    elif dialog_type == "edge":
        # Edge浏览器的打印对话框处理
        pass
    else:
        # 默认处理逻辑
        pass
```

## 注意事项

1. **坐标准确性**：确保Python脚本中的坐标与实际界面元素位置一致
2. **屏幕分辨率**：不同屏幕分辨率可能需要调整坐标
3. **界面变化**：如果网页界面发生变化，需要更新相应的选择器或坐标
4. **权限问题**：确保程序有足够的权限执行鼠标键盘操作

## 依赖项

- .NET 8.0 或更高版本
- Python 3.9 或更高版本
- pyautogui 库
- 其他Python依赖（见requirements_mouse_keyboard.txt）
