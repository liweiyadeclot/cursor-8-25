# Python脚本执行器 (PythonScriptExecutor)

## 概述

`PythonScriptExecutor` 是一个用于在C#应用程序中执行Python脚本的类。它提供了自动查找Python解释器、执行Python脚本、处理输出和错误等功能。

## 主要功能

### 1. 自动Python解释器查找
- 自动在系统中查找Python解释器
- 支持多个Python版本 (3.9, 3.10, 3.11, 3.12)
- 支持常见的Python安装路径
- 支持PATH环境变量中的Python

### 2. Python脚本执行
- 异步执行Python脚本
- 同步执行Python脚本
- 支持脚本参数传递
- 支持工作目录设置
- 支持超时设置

### 3. 输出处理
- 实时捕获Python脚本的标准输出
- 实时捕获Python脚本的错误输出
- 返回详细的执行结果

### 4. 专用方法
- `ExecuteAutoClickerAsync()` - 执行自动化点击流程
- `ExecuteMouseKeyboardAutomationAsync()` - 执行鼠标键盘自动化脚本
- `ExecuteLoginAutomationAsync()` - 执行登录自动化脚本

## 使用方法

### 基本用法

```csharp
// 创建Python脚本执行器实例
var executor = new PythonScriptExecutor();

// 异步执行Python脚本
var result = await executor.ExecuteScriptAsync("script.py", "arg1 arg2");

// 检查执行结果
if (result.Success)
{
    Console.WriteLine($"执行成功: {result.Output}");
}
else
{
    Console.WriteLine($"执行失败: {result.Error}");
}
```

### 同步执行

```csharp
var executor = new PythonScriptExecutor();
var result = executor.ExecuteScript("script.py");
```

### 指定Python路径

```csharp
var executor = new PythonScriptExecutor(@"C:\Python311\python.exe");
```

### 执行专用脚本

```csharp
var executor = new PythonScriptExecutor();

// 执行鼠标键盘自动化
var success = await executor.ExecuteMouseKeyboardAutomationAsync();

// 执行登录自动化
var loginSuccess = await executor.ExecuteLoginAutomationAsync();
```

## 类结构

### PythonScriptExecutor

#### 构造函数
- `PythonScriptExecutor(string pythonPath = null)` - 创建实例，可选指定Python路径

#### 主要方法
- `ExecuteScriptAsync()` - 异步执行Python脚本
- `ExecuteScript()` - 同步执行Python脚本
- `ExecuteAutoClickerAsync()` - 执行自动化点击流程
- `ExecuteMouseKeyboardAutomationAsync()` - 执行鼠标键盘自动化
- `ExecuteLoginAutomationAsync()` - 执行登录自动化

### PythonExecutionResult

#### 属性
- `Success` - 执行是否成功
- `Output` - 标准输出内容
- `Error` - 错误输出内容
- `ExitCode` - 退出代码

## 错误处理

### 常见错误
1. **未找到Python解释器**
   - 确保系统已安装Python
   - 检查Python是否在PATH环境变量中
   - 手动指定Python路径

2. **脚本文件不存在**
   - 检查脚本文件路径是否正确
   - 确保脚本文件存在

3. **执行超时**
   - 增加超时时间
   - 检查脚本是否有无限循环

4. **权限问题**
   - 确保有执行Python脚本的权限
   - 检查工作目录权限

## 测试

运行测试程序：

```bash
# 使用批处理文件
test_python_executor.bat

# 或直接使用dotnet
dotnet run --project "Auto Finan.csproj" -- TestPythonExecutor
```

## 注意事项

1. **Python版本兼容性**
   - 支持Python 3.9及以上版本
   - 建议使用Python 3.10或更高版本

2. **路径处理**
   - 支持相对路径和绝对路径
   - 自动处理路径中的空格和特殊字符

3. **输出编码**
   - 支持中文输出
   - 使用UTF-8编码

4. **性能考虑**
   - 异步执行避免阻塞主线程
   - 可设置超时时间防止无限等待

## 示例项目

项目包含以下示例：
- `TestPythonExecutor.cs` - 测试类
- `test_python_executor.bat` - 测试批处理文件
- 各种Python脚本文件

## 依赖项

- .NET 8.0 或更高版本
- Python 3.9 或更高版本
- 相关Python包（根据具体脚本需求）






