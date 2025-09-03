# 打印按钮Python脚本集成功能说明

## 功能概述

本程序在点击打印报销单按钮后，会自动执行`mouse_keyboard_automation.py`脚本，实现鼠标键盘自动化操作。

## 触发条件

当Excel配置文件中的按钮标题为以下任一值时，会触发Python脚本执行：
- "打印报销单"
- "打印确认单" 
- "报销确认单"

## 执行流程

1. **按钮查找**：程序会查找并点击打印按钮
2. **等待窗口**：等待打印窗口打开（3秒）
3. **执行Python**：自动执行`mouse_keyboard_automation.py`脚本
4. **监控输出**：实时显示Python脚本的输出和错误信息

## Python脚本要求

### 文件位置
程序会按以下顺序查找Python脚本：
1. 当前目录：`mouse_keyboard_automation.py`
2. 上级目录：`../mouse_keyboard_automation.py`
3. 上上级目录：`../../mouse_keyboard_automation.py`
4. 上上上级目录：`../../../mouse_keyboard_automation.py`

### 环境要求
- Python已安装并添加到系统PATH
- 脚本具有执行权限
- 脚本依赖的库已安装

## 执行参数

- **超时时间**：60秒（如果超时会强制终止）
- **工作目录**：Python脚本所在目录
- **输出重定向**：实时显示Python脚本的输出

## 日志输出

程序会显示以下信息：
- 🖨️ 开始处理打印按钮
- ✅ 按钮点击成功
- 🐍 开始执行Python脚本
- 🚀 Python脚本启动信息
- ⏳ 执行进度
- ✅ 执行完成或❌ 执行失败

## 错误处理

- **脚本不存在**：显示错误信息并继续执行
- **Python未安装**：显示错误信息
- **执行超时**：强制终止进程
- **其他错误**：显示详细错误信息

## 使用方法

1. 确保`mouse_keyboard_automation.py`文件存在
2. 在Excel中设置按钮标题为"打印报销单"
3. 运行程序，会自动触发Python脚本执行

## 测试

运行`test_python_integration.bat`来测试Python集成功能。

## 注意事项

1. Python脚本应该能够独立运行
2. 脚本执行时间不应超过60秒
3. 确保Python环境配置正确
4. 脚本输出会实时显示在控制台







