# 自动打印处理功能使用说明

## 概述

本功能实现了在点击报销确认单按钮后，自动等待2秒然后执行Python脚本来处理打印对话框，实现完全自动化的报销流程。

## 功能特性

### 1. 自动化流程
- 运行完整的报销自动化流程
- 自动检查Python环境
- 等待用户点击报销确认单按钮
- 自动等待2秒后处理打印对话框
- 自动保存PDF文件

### 2. 智能处理
- 如果Python环境异常，会跳过打印处理但继续报销流程
- 自动生成带时间戳的文件名
- 详细的错误提示和调试信息

## 使用方法

### 方法1：使用完整自动化流程（推荐）

选择选项4，运行完整的报销自动化流程：

```bash
# 运行程序后选择选项4
4. 完整报销自动化（包含Python打印处理）
```

流程说明：
1. 检查Python环境
2. 运行报销自动化流程
3. 提示用户点击报销确认单按钮
4. 等待2秒
5. 自动处理打印对话框
6. 保存PDF文件

### 方法2：仅执行打印对话框处理

选择选项3，仅执行打印对话框处理：

```bash
# 运行程序后选择选项3
3. 自动执行Python脚本处理打印对话框
```

使用场景：
- 当您已经完成了报销流程，只需要处理打印对话框时
- 用于测试打印对话框处理功能

## 工作流程

### 完整流程示例

```
=== 检查Python环境 ===
✓ Python环境正常

=== 开始报销自动化流程 ===
[报销自动化流程执行...]

=== 准备处理打印对话框 ===
报销流程已完成，请点击报销确认单按钮...
程序将在2秒后自动处理打印对话框...
等待2秒...
开始自动处理打印对话框...

✓ 打印对话框处理成功！
文件已保存到: C:\Users\FH\PycharmProjects\CursorCode8-5\pdf_output\报销单_20241201_143022.pdf
```

### 错误处理示例

```
=== 检查Python环境 ===
警告：Python环境异常，将跳过打印对话框自动处理
请确保已安装Python和pyautogui库

=== 开始报销自动化流程 ===
[报销自动化流程正常执行，但跳过打印处理]
```

## 环境要求

### Python环境
1. 安装Python 3.7+
2. 安装pyautogui库：
   ```bash
   pip install pyautogui
   ```

### 配置文件
确保`config.py`中的坐标配置正确：
```python
PRINT_DIALOG_COORDINATES = {
    "print_button": {"x": 1562, "y": 1083},
    "filepath_input": {"x": 720, "y": 98},
    "filename_input": {"x": 404, "y": 17},
    "save_button": {"x": 971, "y": 869},
    "yes_button": {"x": 700, "y": 450}
}
PRINT_FILE_PATH = r"C:\Users\FH\PycharmProjects\CursorCode8-5\pdf_output"
```

## 文件保存规则

### 文件名格式
```
报销单_YYYYMMDD_HHMMSS.pdf
```

示例：
- `报销单_20241201_143022.pdf`
- `报销单_20241201_150530.pdf`

### 保存路径
默认保存到配置文件中的`PRINT_FILE_PATH`路径。

## 故障排除

### 常见问题

1. **Python环境异常**
   ```
   警告：Python环境异常，将跳过打印对话框自动处理
   ```
   解决：
   - 检查Python是否正确安装
   - 运行 `pip install pyautogui`
   - 确保Python在系统PATH中

2. **打印对话框处理失败**
   ```
   ✗ 打印对话框处理失败
   ```
   解决：
   - 检查打印对话框是否正确显示
   - 验证坐标配置是否正确
   - 确认文件路径是否存在

3. **坐标不准确**
   ```
   点击鼠标失败
   ```
   解决：
   - 使用选项2中的坐标获取功能重新获取坐标
   - 更新`config.py`中的坐标配置

### 调试技巧

1. **获取鼠标位置**
   ```csharp
   var automation = new ReimbursementAutomationWithPython();
   await automation.GetMousePosition();
   ```

2. **手动触发打印处理**
   ```csharp
   var automation = new ReimbursementAutomationWithPython();
   await automation.ManualTriggerPrintDialog();
   ```

3. **检查Python脚本**
   ```bash
   python mouse_keyboard_automation.py --check
   ```

## 配置说明

### 坐标配置
在`config.py`中配置打印对话框的坐标：

```python
PRINT_DIALOG_COORDINATES = {
    "print_button": {"x": 1562, "y": 1083},      # 打印按钮
    "filepath_input": {"x": 720, "y": 98},       # 文件路径输入框
    "filename_input": {"x": 404, "y": 17},       # 文件名输入框
    "save_button": {"x": 971, "y": 869},         # 保存按钮
    "yes_button": {"x": 700, "y": 450}           # 确认按钮（可选）
}
```

### 文件路径配置
```python
PRINT_FILE_PATH = r"C:\Users\FH\PycharmProjects\CursorCode8-5\pdf_output"
```

## 使用建议

### 1. 首次使用
1. 先选择选项2测试Python环境
2. 使用选项3测试打印对话框处理
3. 确认无误后使用选项4运行完整流程

### 2. 日常使用
直接使用选项4运行完整自动化流程

### 3. 调试时
- 使用选项2获取鼠标位置
- 使用选项3单独测试打印处理
- 检查日志文件了解详细错误信息

## 扩展功能

### 1. 自定义等待时间
可以修改代码中的等待时间：
```csharp
// 修改等待时间（毫秒）
await Task.Delay(3000); // 等待3秒
```

### 2. 自定义文件路径
可以修改保存路径：
```csharp
string filePath = @"C:\自定义路径\pdf_output";
```

### 3. 自定义文件名格式
可以修改文件名格式：
```csharp
string fileName = $"自定义前缀_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
```

## 总结

自动打印处理功能实现了报销流程的完全自动化，从数据填写到PDF保存，整个过程无需人工干预。通过合理的错误处理和用户提示，确保了流程的稳定性和用户体验。
