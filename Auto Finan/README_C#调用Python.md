# C#调用Python脚本使用说明

## 概述

本文档介绍如何在C#项目中调用Python脚本，实现鼠标键盘自动化功能。通过`PythonInterop`类，您可以在C#代码中轻松调用Python的`mouse_keyboard_automation.py`脚本。

## 实现方式

### 方法1：使用Process调用Python脚本（推荐）

这是最简单直接的方法，通过C#的Process类来执行Python脚本。

## 文件结构

```
Auto Finan/
├── Program.cs                    # C#主程序
├── PythonInterop.cs             # Python脚本调用器
├── mouse_keyboard_automation.py # Python鼠标键盘自动化脚本
├── config.py                    # Python配置文件
└── requirements_mouse_keyboard.txt # Python依赖包
```

## 使用方法

### 1. 基本使用

```csharp
using AutoFinan;

// 创建Python调用器
var pythonInterop = new PythonInterop();

// 检查Python环境
bool envOk = await pythonInterop.CheckPythonEnvironmentAsync();
if (envOk)
{
    Console.WriteLine("Python环境正常");
}
```

### 2. 获取鼠标位置

```csharp
// 获取当前鼠标位置
string mousePos = await pythonInterop.GetMousePositionAsync();
Console.WriteLine($"鼠标位置: {mousePos}");
```

### 3. 执行文件保存流程

```csharp
// 定义坐标配置
var coordinates = new
{
    button = new { x = 1562, y = 1083 },
    filepath_input = new { x = 720, y = 98 },
    filename_input = new { x = 404, y = 17 },
    save_button = new { x = 971, y = 869 }
};

// 序列化为JSON
string coordinatesJson = System.Text.Json.JsonSerializer.Serialize(coordinates);

// 执行文件保存流程
bool success = await pythonInterop.ExecuteFileSaveProcessAsync(
    @"C:\Users\FH\Documents\test",
    "test_file.pdf",
    coordinatesJson
);

Console.WriteLine($"执行结果: {(success ? "成功" : "失败")}");
```

### 4. 执行打印对话框处理

```csharp
// 执行打印对话框处理
bool success = await pythonInterop.ExecutePrintDialogProcessAsync(
    @"C:\Users\FH\Documents\pdf_output",
    "test_print.pdf"
);

Console.WriteLine($"执行结果: {(success ? "成功" : "失败")}");
```

### 5. 自定义Python路径

```csharp
// 如果Python不在PATH中，可以指定完整路径
var pythonInterop = new PythonInterop(
    pythonPath: @"C:\Python39\python.exe",
    scriptPath: @"C:\path\to\mouse_keyboard_automation.py"
);
```

## 完整示例

```csharp
using System;
using System.Threading.Tasks;
using AutoFinan;

class Program
{
    static async Task Main(string[] args)
    {
        Console.WriteLine("=== C#调用Python脚本示例 ===");
        
        var pythonInterop = new PythonInterop();
        
        try
        {
            // 1. 检查Python环境
            Console.WriteLine("检查Python环境...");
            bool envOk = await pythonInterop.CheckPythonEnvironmentAsync();
            if (!envOk)
            {
                Console.WriteLine("Python环境异常，请检查安装");
                return;
            }
            
            // 2. 获取鼠标位置
            Console.WriteLine("获取鼠标位置...");
            string mousePos = await pythonInterop.GetMousePositionAsync();
            Console.WriteLine($"鼠标位置: {mousePos}");
            
            // 3. 执行文件保存流程
            Console.WriteLine("执行文件保存流程...");
            var coordinates = new
            {
                button = new { x = 1562, y = 1083 },
                filepath_input = new { x = 720, y = 98 },
                filename_input = new { x = 404, y = 17 },
                save_button = new { x = 971, y = 869 }
            };
            
            string coordinatesJson = System.Text.Json.JsonSerializer.Serialize(coordinates);
            bool fileSaveSuccess = await pythonInterop.ExecuteFileSaveProcessAsync(
                @"C:\Users\FH\Documents\test",
                "test_file.pdf",
                coordinatesJson
            );
            
            Console.WriteLine($"文件保存: {(fileSaveSuccess ? "成功" : "失败")}");
            
            // 4. 执行打印对话框处理
            Console.WriteLine("执行打印对话框处理...");
            bool printDialogSuccess = await pythonInterop.ExecutePrintDialogProcessAsync(
                @"C:\Users\FH\Documents\pdf_output",
                "test_print.pdf"
            );
            
            Console.WriteLine($"打印对话框: {(printDialogSuccess ? "成功" : "失败")}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"执行出错: {ex.Message}");
        }
    }
}
```

## Python脚本命令行参数

Python脚本支持以下命令行参数：

### 基本参数
- `--check`: 检查Python环境
- `--demo`: 运行演示
- `--config`: 使用配置文件执行

### 操作参数
- `--operation <操作类型>`: 指定操作类型
  - `get_position`: 获取鼠标位置
  - `file_save`: 执行文件保存流程
  - `print_dialog`: 执行打印对话框处理

### 文件操作参数
- `--filepath <文件路径>`: 指定文件路径
- `--filename <文件名>`: 指定文件名
- `--coordinates <JSON坐标>`: 指定坐标配置（JSON格式）

### 示例命令

```bash
# 检查环境
python mouse_keyboard_automation.py --check

# 获取鼠标位置
python mouse_keyboard_automation.py --operation get_position

# 执行文件保存流程
python mouse_keyboard_automation.py --operation file_save --filepath "C:\test" --filename "test.pdf" --coordinates '{"button":{"x":100,"y":100}}'

# 执行打印对话框处理
python mouse_keyboard_automation.py --operation print_dialog --filepath "C:\output" --filename "test.pdf"

# 运行演示
python mouse_keyboard_automation.py --demo
```

## 环境要求

### Python环境
1. 安装Python 3.7+
2. 安装pyautogui库：
   ```bash
   pip install pyautogui
   ```

### C#环境
1. .NET 6.0+
2. 确保Python在系统PATH中，或者指定完整路径

## 错误处理

### 常见问题

1. **Python未找到**
   ```
   错误：未找到Python环境
   解决：确保Python已安装并在PATH中，或指定完整路径
   ```

2. **pyautogui未安装**
   ```
   错误：No module named 'pyautogui'
   解决：运行 pip install pyautogui
   ```

3. **脚本路径错误**
   ```
   错误：找不到脚本文件
   解决：确保mouse_keyboard_automation.py文件存在
   ```

4. **权限问题**
   ```
   错误：访问被拒绝
   解决：以管理员身份运行程序
   ```

### 调试技巧

```csharp
// 启用详细日志
var pythonInterop = new PythonInterop();
var result = await pythonInterop.ExecutePythonScriptAsync("--check");
Console.WriteLine($"Python输出: {result}");
```

## 性能优化

### 1. 缓存Python进程
```csharp
// 避免重复启动Python进程
private static PythonInterop _pythonInterop;

public static async Task<bool> ExecuteOperationAsync()
{
    if (_pythonInterop == null)
    {
        _pythonInterop = new PythonInterop();
    }
    
    return await _pythonInterop.ExecuteMouseKeyboardOperationAsync("file_save");
}
```

### 2. 批量操作
```csharp
// 批量执行多个操作
var tasks = new List<Task<bool>>();
for (int i = 0; i < 10; i++)
{
    tasks.Add(pythonInterop.ExecutePrintDialogProcessAsync(
        @"C:\output", 
        $"file_{i}.pdf"
    ));
}

var results = await Task.WhenAll(tasks);
```

## 安全考虑

1. **路径验证**: 确保文件路径安全
2. **权限检查**: 验证程序有足够权限
3. **异常处理**: 捕获并处理所有异常
4. **超时设置**: 避免无限等待

```csharp
// 安全示例
public async Task<bool> SafeExecuteAsync(string filePath)
{
    // 验证路径
    if (!Path.IsPathRooted(filePath))
    {
        throw new ArgumentException("路径必须是绝对路径");
    }
    
    // 设置超时
    using var cts = new CancellationTokenSource(TimeSpan.FromMinutes(5));
    
    try
    {
        return await pythonInterop.ExecuteFileSaveProcessAsync(
            filePath, "test.pdf", coordinatesJson
        ).WaitAsync(cts.Token);
    }
    catch (OperationCanceledException)
    {
        Console.WriteLine("操作超时");
        return false;
    }
}
```

## 扩展功能

### 1. 配置文件支持
```csharp
public class PythonConfig
{
    public string PythonPath { get; set; } = "python";
    public string ScriptPath { get; set; } = "mouse_keyboard_automation.py";
    public int TimeoutSeconds { get; set; } = 300;
}
```

### 2. 日志记录
```csharp
public class PythonInteropWithLogging : PythonInterop
{
    private readonly ILogger _logger;
    
    public PythonInteropWithLogging(ILogger logger)
    {
        _logger = logger;
    }
    
    public override async Task<string> ExecutePythonScriptAsync(string arguments)
    {
        _logger.LogInformation($"执行Python脚本: {arguments}");
        var result = await base.ExecutePythonScriptAsync(arguments);
        _logger.LogInformation($"Python脚本执行完成: {result}");
        return result;
    }
}
```

## 总结

通过`PythonInterop`类，您可以轻松在C#中调用Python脚本，实现复杂的鼠标键盘自动化功能。这种方法结合了C#的强类型特性和Python的自动化库优势，为您的项目提供了灵活的解决方案。
