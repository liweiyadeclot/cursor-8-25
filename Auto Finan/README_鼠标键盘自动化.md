# 鼠标键盘自动化功能使用说明

## 概述

本项目实现了基于C#的鼠标键盘输入模拟功能，参考了Python版本的`login_automation.py`中的实现思路，使用`WindowsInput`库来实现Windows系统下的鼠标键盘自动化操作。

## 功能特性

### 1. 基本操作
- **鼠标点击**: 模拟鼠标在指定坐标的点击操作
- **键盘输入**: 模拟键盘文本输入
- **按键操作**: 模拟单个按键或组合键操作

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

确保已安装`WindowsInput`库：

```bash
dotnet add package WindowsInput
```

## 使用方法

### 1. 基本使用

```csharp
var automation = new MouseKeyboardAutomation();

// 鼠标点击
automation.ClickMouse(100, 100, 500);

// 键盘输入
automation.TypeText("Hello World", 500);

// 按键操作
automation.PressKey(VirtualKeyCode.RETURN, 500);

// 组合键
automation.PressCombination(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_S);
```

### 2. 文件保存流程

```csharp
var coordinates = new FileSaveCoordinates
{
    ButtonX = 1562,
    ButtonY = 1083,
    FilePathInputX = 720,
    FilePathInputY = 98,
    FileNameInputX = 404,
    FileNameInputY = 17,
    SaveButtonX = 971,
    SaveButtonY = 869
};

await automation.ExecuteFileSaveProcess(
    @"C:\Users\FH\Documents\test", 
    "test_file.pdf", 
    coordinates
);
```

### 3. 打印对话框处理

```csharp
var coordinates = new PrintDialogCoordinates
{
    PrintButtonX = 1562,
    PrintButtonY = 1083,
    FilePathInputX = 720,
    FilePathInputY = 98,
    FileNameInputX = 404,
    FileNameInputY = 17,
    SaveButtonX = 971,
    SaveButtonY = 869,
    YesButtonX = 700,
    YesButtonY = 450
};

await automation.ExecutePrintDialogProcess(
    @"C:\Users\FH\Documents\pdf_output", 
    "报销单.pdf", 
    coordinates
);
```

## 坐标获取方法

### 1. 使用Windows自带的坐标获取工具
- 按`Win + R`，输入`cmd`
- 在命令行中输入：`powershell -command "Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Cursor]::Position"`
- 移动鼠标到目标位置，按回车获取坐标

### 2. 使用第三方工具
- **AutoHotkey Window Spy**: 可以实时显示鼠标坐标
- **Screen Ruler**: 屏幕标尺工具
- **MousePos**: 专门的鼠标坐标获取工具

## 配置说明

### 坐标配置类

#### FileSaveCoordinates
```csharp
public class FileSaveCoordinates
{
    public int ButtonX { get; set; }        // 按钮X坐标
    public int ButtonY { get; set; }        // 按钮Y坐标
    public int FilePathInputX { get; set; } // 文件路径输入框X坐标
    public int FilePathInputY { get; set; } // 文件路径输入框Y坐标
    public int FileNameInputX { get; set; } // 文件名输入框X坐标
    public int FileNameInputY { get; set; } // 文件名输入框Y坐标
    public int SaveButtonX { get; set; }    // 保存按钮X坐标
    public int SaveButtonY { get; set; }    // 保存按钮Y坐标
}
```

#### PrintDialogCoordinates
```csharp
public class PrintDialogCoordinates
{
    public int PrintButtonX { get; set; }   // 打印按钮X坐标
    public int PrintButtonY { get; set; }   // 打印按钮Y坐标
    public int FilePathInputX { get; set; } // 文件路径输入框X坐标
    public int FilePathInputY { get; set; } // 文件路径输入框Y坐标
    public int FileNameInputX { get; set; } // 文件名输入框X坐标
    public int FileNameInputY { get; set; } // 文件名输入框Y坐标
    public int SaveButtonX { get; set; }    // 保存按钮X坐标
    public int SaveButtonY { get; set; }    // 保存按钮Y坐标
    public int YesButtonX { get; set; }     // 确认按钮X坐标（可选）
    public int YesButtonY { get; set; }     // 确认按钮Y坐标（可选）
}
```

## 注意事项

1. **坐标准确性**: 坐标值必须准确，建议多次测试确认
2. **屏幕分辨率**: 不同屏幕分辨率下坐标可能不同
3. **窗口位置**: 确保目标窗口在前台且位置固定
4. **权限要求**: 某些操作可能需要管理员权限
5. **防病毒软件**: 某些防病毒软件可能会阻止模拟输入

## 错误处理

代码包含完整的异常处理机制：
- 坐标无效时的错误提示
- 操作失败时的重试机制
- 详细的日志输出

## 与Python版本的对比

| 功能 | Python版本 | C#版本 |
|------|------------|--------|
| 鼠标点击 | pyautogui | WindowsInput |
| 键盘输入 | pyautogui | WindowsInput |
| 坐标获取 | pyautogui.position() | 手动获取 |
| 错误处理 | try-except | try-catch |
| 日志记录 | logging | Console.WriteLine |

## 扩展功能

可以根据需要扩展以下功能：
- 图像识别定位
- 多显示器支持
- 更复杂的操作序列
- 配置文件支持
- GUI界面

## 示例程序

运行程序后选择相应的功能：
1. 财务报销自动化
2. 鼠标键盘自动化演示
3. 打印对话框处理

选择选项2可以查看完整的演示流程。
