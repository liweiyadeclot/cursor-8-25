# Python调用方式说明

## 概述
本文档说明了如何在C#程序中调用Python脚本，以及如何测试和调试这种调用方式。

## 调用方式

### 1. 基本调用格式
```bash
python test_mouse_keyboard.py --config config.json --folder "C:\Users\FH\Documents\报销单" --file "test_file.pdf"
```

### 2. C#程序中的调用
在C#程序中，通过`PythonScriptExecutor`类来调用Python脚本：

```csharp
var result = await pythonExecutor.ExecuteScriptAsync(
    scriptPath: "test_mouse_keyboard.py",
    arguments: "--config config.json --folder \"C:\\Users\\FH\\Documents\\报销单\" --file \"报销单_{DateTime.Now:yyyyMMdd_HHmmss}.pdf\"",
    timeoutMilliseconds: 120000 // 2分钟超时
);
```

### 3. 参数说明
- `--config`: 配置文件路径（默认: config.json）
- `--folder`: 要输入的文件夹路径
- `--file`: 要输入的文件名
- `--delay`: 每次点击后的延迟时间（可选，默认: 1.0秒）

## 测试工具

### 1. 直接调用测试
```bash
test_direct_python_call.bat
```
这个工具会测试两种调用方式：
- 批处理直接调用
- Python脚本内部调用

### 2. 简化调用测试
```bash
test_python_call_simple.bat
```
这个工具测试基本的Python调用功能。

### 3. 诊断工具
```bash
diagnose_python_issue.bat
```
这个工具会检查Python环境、文件存在性、包安装等。

## 常见问题

### 1. Python脚本执行失败
**症状**: 错误信息为空
**可能原因**:
- Python环境问题
- 文件路径问题
- 权限问题

**解决方法**:
1. 运行 `diagnose_python_issue.bat` 进行诊断
2. 检查Python是否正确安装
3. 检查必要文件是否存在
4. 检查pyautogui是否安装

### 2. 配置文件读取失败
**症状**: "local variable 'json' referenced before assignment"
**解决方法**: 修复变量名冲突

### 3. 超时问题
**症状**: 执行超时
**解决方法**: 增加超时时间或检查脚本执行时间

## 调试步骤

1. **基础检查**:
   ```bash
   python --version
   pip list | findstr pyautogui
   ```

2. **文件检查**:
   ```bash
   dir test_mouse_keyboard.py
   dir config.json
   ```

3. **直接测试**:
   ```bash
   python test_mouse_keyboard.py --help
   ```

4. **完整测试**:
   ```bash
   test_direct_python_call.bat
   ```

## 最佳实践

1. **使用绝对路径**: 确保文件路径正确
2. **检查工作目录**: 确保在正确的目录中执行
3. **添加错误处理**: 在C#代码中添加完善的错误处理
4. **日志记录**: 记录详细的执行日志
5. **超时设置**: 设置合理的超时时间

## 示例输出

### 成功调用
```
=== 简化Python调用测试 ===

命令行参数:
  0: test_python_call_simple.py
  1: --config
  2: config.json
  3: --folder
  4: C:\Users\FH\Documents\报销单
  5: --file
  6: test_file.pdf

当前工作目录: C:\Users\FH\source\repos\Auto Finan\Auto Finan

✓ config.json 存在
✓ test_mouse_keyboard.py 存在

✓ 简化测试成功完成
退出码: 0
```

### 失败调用
```
✗ Python环境测试失败
错误: 无错误信息
请先运行 debug_python_execution.bat 进行诊断
```






