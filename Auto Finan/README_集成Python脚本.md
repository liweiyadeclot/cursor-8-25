# 报销自动化流程集成Python脚本使用说明

## 概述

本功能在原有的财务报销自动化流程中集成了Python脚本调用，当点击报销确认单按钮后，会自动等待2秒然后执行Python脚本来处理打印对话框，实现完全自动化的报销流程。

## 功能特性

### 1. 无缝集成
- 在原有的报销自动化流程中直接集成Python脚本
- 无需修改现有的报销流程逻辑
- 自动处理打印对话框，无需人工干预

### 2. 智能处理
- 自动检查Python环境
- 如果Python脚本执行失败，会自动回退到原有的处理方式
- 自动生成带时间戳的文件名
- 详细的日志记录

### 3. 错误处理
- 如果Python环境不可用，会使用备用方案
- 如果Python脚本执行失败，会回退到原有的坐标点击方式
- 完整的异常处理和日志记录

## 工作流程

### 完整流程
```
1. 运行报销自动化流程
2. 填写报销信息
3. 点击报销确认单按钮
4. 等待2秒（确保打印页面加载）
5. 自动执行Python脚本处理打印对话框
6. 保存PDF文件
7. 完成报销流程
```

### 详细执行步骤
```
=== 报销自动化流程 ===
[填写报销信息...]
[点击报销确认单按钮]

查找网页上的打印确认单按钮...
✓ 网页打印确认单按钮点击成功
等待2秒钟，确保Chrome打印页面加载完成...
开始自动执行Python脚本处理打印对话框...

项目号: 12345, 总金额: 1000.00
准备保存文件: 报销单_20241201_143022.pdf
保存路径: C:\Users\FH\PycharmProjects\CursorCode8-5\pdf_output
执行Python脚本命令: python mouse_keyboard_automation.py --operation print_dialog --filepath "C:\Users\FH\PycharmProjects\CursorCode8-5\pdf_output" --filename "报销单_20241201_143022.pdf"
Python脚本执行成功: ✓ 打印对话框处理执行成功
✓ Python脚本处理打印对话框成功
文件已保存到: C:\Users\FH\PycharmProjects\CursorCode8-5\pdf_output/报销单_20241201_143022.pdf
```

## 技术实现

### 1. 修改的方法
- `click_print_button()`: 主要方法，负责点击打印按钮并调用Python脚本
- `auto_execute_python_print_dialog()`: 自动执行Python脚本的方法
- `_execute_python_print_script()`: 实际执行Python脚本的方法

### 2. Python脚本调用
```python
# 构建命令
command = [
    sys.executable,  # 使用当前Python解释器
    "mouse_keyboard_automation.py",
    "--operation", "print_dialog",
    "--filepath", file_path,
    "--filename", file_name
]

# 异步执行
process = await asyncio.create_subprocess_exec(
    *command,
    stdout=asyncio.subprocess.PIPE,
    stderr=asyncio.subprocess.PIPE
)
```

### 3. 错误处理机制
```python
try:
    # 尝试执行Python脚本
    success = await self._execute_python_print_script(file_path, file_name)
    if success:
        logger.info("✓ Python脚本处理成功")
    else:
        # 回退到备用方案
        await self._handle_print_dialog_fallback()
except Exception as e:
    # 异常时使用备用方案
    await self._handle_print_dialog_fallback()
```

## 环境要求

### Python环境
1. 安装Python 3.7+
2. 安装pyautogui库：
   ```bash
   pip install pyautogui
   ```

### 文件要求
- `mouse_keyboard_automation.py`: Python脚本文件
- `config.py`: 配置文件，包含坐标和路径配置

### 配置文件
确保`config.py`中的配置正确：
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

## 使用方法

### 1. 正常运行
直接运行报销自动化流程，当遇到打印按钮时会自动处理：
```bash
python login_automation.py
```

### 2. 测试Python脚本
单独测试Python脚本功能：
```bash
python mouse_keyboard_automation.py --operation print_dialog --filepath "C:\test" --filename "test.pdf"
```

### 3. 检查环境
检查Python环境是否正常：
```bash
python mouse_keyboard_automation.py --check
```

## 故障排除

### 常见问题

1. **Python脚本执行失败**
   ```
   Python脚本执行失败: [错误信息]
   ```
   解决：
   - 检查Python是否正确安装
   - 确认pyautogui库已安装
   - 验证脚本文件路径是否正确

2. **备用方案自动启用**
   ```
   ⚠ Python脚本处理失败，尝试备用方案
   ```
   说明：系统自动回退到原有的坐标点击方式，流程仍可正常完成

3. **文件保存失败**
   ```
   文件已保存到: [路径]
   ```
   检查：
   - 文件路径是否存在
   - 是否有写入权限
   - 磁盘空间是否充足

### 调试技巧

1. **查看详细日志**
   ```python
   # 在代码中启用详细日志
   logging.getLogger().setLevel(logging.DEBUG)
   ```

2. **手动测试Python脚本**
   ```bash
   python mouse_keyboard_automation.py --demo
   ```

3. **检查坐标配置**
   ```bash
   python mouse_keyboard_automation.py --operation get_position
   ```

## 性能优化

### 1. 异步执行
Python脚本使用异步方式执行，不会阻塞主流程：
```python
process = await asyncio.create_subprocess_exec(...)
```

### 2. 超时控制
可以添加超时控制避免无限等待：
```python
try:
    stdout, stderr = await asyncio.wait_for(
        process.communicate(), 
        timeout=30.0
    )
except asyncio.TimeoutError:
    process.kill()
    return False
```

### 3. 错误恢复
如果Python脚本失败，自动回退到备用方案，确保流程不中断。

## 扩展功能

### 1. 自定义等待时间
可以修改等待时间：
```python
# 修改等待时间（秒）
await asyncio.sleep(3)  # 等待3秒
```

### 2. 自定义文件名格式
可以修改文件名生成规则：
```python
# 自定义文件名格式
file_name = f"自定义前缀_{time.strftime('%Y%m%d_%H%M%S')}.pdf"
```

### 3. 添加更多Python脚本功能
可以在Python脚本中添加更多功能：
- 图像识别
- 多显示器支持
- 更复杂的操作序列

## 总结

通过集成Python脚本，报销自动化流程实现了完全自动化，从数据填写到PDF保存，整个过程无需人工干预。系统具备智能的错误处理和回退机制，确保流程的稳定性和可靠性。
