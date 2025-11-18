# Excel财务自动化系统 - 完整使用流程

## 📋 系统概述

本系统实现了从Excel财务数据到Playwright MCP自动化操作的完整流程，支持报销业务、差旅业务和劳务业务三种类型。

---

## 🚀 完整工作流程

### **第1步：准备Excel数据**

在Excel文件中准备财务数据，必需包含以下列：
- `序号` - 用于分组数据
- `账号` / `登录界面工号` - 登录用户名
- `密码` / `登录界面密码` - 登录密码
- `业务大类` / `选择业务大类` - 业务类型（报销业务、业务出差旅费、酬金业务）
- `!已生成MCP提示词` - 标记列（程序会自动填写）

---

### **第2步：批量生成MCP提示词**

运行批量处理器：

```bash
cd "C:\Users\FH\source\repos\Auto Finan\Auto Finan\LLM_Integration"
python excel_batch_processor.py
```

**交互过程：**
```
=== Excel批量处理器 - 生成MCP提示词 ===
Excel文件路径: C:\Users\FH\source\repos\Auto Finan\Auto Finan\420财务050823.xlsx
输出目录: C:\Users\FH\source\repos\Auto Finan\Auto Finan\LLM_Integration\mcp_prompts

请输入工作表名称（如：3-报销、3-差旅、3-劳务）：3-报销
```

**程序会自动：**
1. ✅ 检查"!已生成MCP提示词"列，跳过已生成的序号
2. ✅ 读取所有未处理序号的数据
3. ✅ 转换为JSON格式
4. ✅ 生成MCP提示词
5. ✅ 保存到文件（格式：`未预约-{项目号}-{总金额}-{时间戳}.txt`）
6. ✅ 每个文件间隔1.5秒（防止覆盖）
7. ✅ 更新Excel的"!已生成MCP提示词"列

**输出示例：**
```
序号: 1
项目号: M112023ZHCG0006
总金额: 100
已保存至: 未预约-M112023ZHCG0006-100-20251009-14-30-25.txt
MCP提示词长度: 845 字符
```

---

### **第3步：执行MCP自动化操作**

#### **方法1：手动执行（在Cursor中）**

1. 打开生成的MCP提示词文件（如 `未预约-M112023ZHCG0006-100-20251009-14-30-25.txt`）
2. 复制文件内容
3. 在Cursor中直接粘贴并发送给AI
4. AI会调用Playwright MCP执行自动化操作

#### **方法2：程序化执行（未来扩展）**

可以编写Python脚本直接调用Playwright API执行。

---

### **第4步：标记为已完成**

执行完MCP操作后，标记文件为已完成：

#### **方法1：使用标记工具**

```bash
python mark_completed_example.py
```

**交互过程：**
```
请输入要标记为已完成的文件名：
未预约-M112023ZHCG0006-100-20251009-14-30-25.txt
```

程序会自动：
1. ✅ 重命名文件：`未预约-xxx.txt` → `已预约-xxx.txt`
2. ✅ 更新Excel的"!已生成MCP提示词"列

#### **方法2：在代码中调用**

```python
from workflow_core import mark_mcp_file_as_completed

new_filename = mark_mcp_file_as_completed(
    excel_path="C:\\path\\to\\file.xlsx",
    sheet_name="3-报销",
    old_filename="未预约-M112023ZHCG0006-100-20251009-14-30-25.txt"
)

print(f"新文件名: {new_filename}")
```

---

## 📁 文件结构

```
LLM_Integration/
├── excel_batch_processor.py      # 主程序：批量生成MCP提示词
├── workflow_core.py               # 核心：工作流程控制和MCP生成
├── excel_to_nl.py                 # Excel数据提取和转换
├── mark_completed_example.py      # 工具：标记文件为已完成
├── field_type_mapping.xlsx        # 配置：字段类型映射
├── mcp_prompts/                   # 输出目录：MCP提示词文件
│   ├── 未预约-M112023ZHCG0006-100-20251009-14-30-25.txt
│   ├── 已预约-M112023ZHCG0006-100-20251009-14-30-27.txt
│   └── ...
└── README_完整使用流程.md         # 本文档
```

---

## 🎯 文件命名规则

### **生成时（未预约）：**
```
未预约-{项目号}-{总金额}-{时间戳}.txt
```

示例：
- `未预约-M112023ZHCG0006-100-20251009-14-30-25.txt`
- `未预约-M112023ZHCG0006-200-20251009-14-30-27.txt`

### **完成后（已预约）：**
```
已预约-{项目号}-{总金额}-{时间戳}.txt
```

示例：
- `已预约-M112023ZHCG0006-100-20251009-14-30-25.txt`
- `已预约-M112023ZHCG0006-200-20251009-14-30-27.txt`

---

## 📊 Excel标记列效果

| 序号 | 项目号 | ... | !已生成MCP提示词 |
|------|--------|-----|------------------|
| 1 | M112023ZHCG0006 | ... | 未预约-M112023ZHCG0006-100-20251009-14-30-25.txt |
| 2 | M112023ZHCG0006 | ... | 已预约-M112023ZHCG0006-100-20251009-14-30-27.txt ✅ |
| 3 | M112023ZHCG0006 | ... | 未预约-M112023ZHCG0006-200-20251009-14-30-28.txt |

---

## 🔧 配置说明

### **修改Excel文件路径**

在 `excel_batch_processor.py` 第132行：
```python
EXCEL_FILE_PATH = r"C:\Users\FH\source\repos\Auto Finan\Auto Finan\420财务050823.xlsx"
```

### **修改输出目录**

在 `excel_batch_processor.py` 第133行：
```python
OUTPUT_DIR = r"C:\Users\FH\source\repos\Auto Finan\Auto Finan\LLM_Integration\mcp_prompts"
```

---

## 🎨 支持的字段类型

在 `field_type_mapping.xlsx` 中配置：

| 标记 | 说明 | 生成格式 |
|------|------|----------|
| `i` | 输入框 | 在{label}输入框中输入{value} |
| `c` | 下拉框 | 在{label}下拉框中选择{value} |
| `r` | 单选按钮 | 点击{label}radio button |
| `d` | 日期选择 | 选择日期{label}为{value} |
| `empty` | 自定义格式 | {label}内容为{value} |
| 空 | 参考信息 | 跳过不生成 |

---

## 💡 使用技巧

### **技巧1：增量处理（智能跳过）**

程序会自动跳过已生成的序号，支持增量处理：
```bash
# 第1次运行：处理序号1、2、3
python excel_batch_processor.py
> 3-报销
# 输出：处理了序号1、2、3

# 添加新数据到Excel（序号4、5）

# 第2次运行：只处理新增的序号
python excel_batch_processor.py
> 3-报销
# 输出：跳过已生成的序号: 1, 2, 3
#      待处理序号: 4, 5
```

### **技巧2：批量处理多个工作表**

可以多次运行程序，分别处理不同的工作表：
```bash
# 第1次运行：处理报销业务
python excel_batch_processor.py
> 3-报销

# 第2次运行：处理差旅业务
python excel_batch_processor.py
> 3-差旅

# 第3次运行：处理劳务业务
python excel_batch_processor.py
> 3-劳务
```

### **技巧3：快速查找未完成的任务**

在Excel中筛选"!已生成MCP提示词"列：
- 包含"未预约" → 已生成但未执行
- 包含"已预约" → 已执行完成
- 空白 → 未生成

### **技巧4：时间戳格式说明**

时间戳格式：`YYYYMMDD-HH-MM-SS`
- `20251009` - 2025年10月9日
- `14-30-25` - 14时30分25秒

---

## 🐛 常见问题

### **Q1：Excel文件被占用无法保存**
**A：** 关闭Excel程序后再运行脚本。

### **Q2：未找到"!已生成MCP提示词"列**
**A：** 在Excel中手动添加此列（列名必须完全一致，包括感叹号）。

### **Q3：文件名中的总金额不正确**
**A：** 检查Excel中的金额数据格式，确保是纯数字。

### **Q4：时间戳相同导致文件覆盖**
**A：** 程序已自动添加1.5秒延迟，正常情况不会覆盖。

---

## 📞 API调用示例

### **在其他Python脚本中使用**

```python
from workflow_core import (
    batch_process_excel_to_mcp_direct,
    mark_mcp_file_as_completed
)

# 1. 批量生成MCP提示词
results = batch_process_excel_to_mcp_direct(
    excel_path="C:\\path\\to\\file.xlsx",
    sheet_name="3-报销"
)

# 2. 执行MCP操作（你的代码）
for result in results:
    mcp_prompt = result['mcp_prompt']
    # ... 执行MCP操作 ...

# 3. 标记为已完成
for result in results:
    filename = f"未预约-{result['json_data']['project']['projectNumber']}-..."
    mark_mcp_file_as_completed(
        excel_path="C:\\path\\to\\file.xlsx",
        sheet_name="3-报销",
        old_filename=filename
    )
```

---

## 🎉 总结

整个系统实现了：
1. ✅ Excel数据自动提取
2. ✅ JSON结构化转换
3. ✅ MCP提示词自动生成
4. ✅ 文件自动保存和命名
5. ✅ Excel自动标记
6. ✅ 完成状态管理

**零LLM依赖，100%确定性，支持大批量数据处理！** 🚀

---

**更新日期：** 2025-10-09  
**版本：** v2.1

