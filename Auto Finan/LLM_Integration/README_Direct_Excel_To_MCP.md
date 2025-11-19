# Excel到MCP提示词直接转换

## 📋 功能概述

本模块实现了从Excel财务数据直接生成Playwright MCP提示词的完整流程，**无需经过LLM自然语言生成环节**，显著提升了性能和稳定性。

## 🎯 核心优势

### 方案对比

| 特性 | 旧方案（Excel→NL→JSON→MCP） | **新方案（Excel→JSON→MCP）** |
|------|---------------------------|------------------------------|
| 速度 | 慢（需调用LLM 2次） | **快（无LLM调用）** |
| 稳定性 | 受上下文长度限制 | **不受限制** |
| 准确性 | 依赖LLM理解 | **100%确定性** |
| 适用场景 | 小批量数据 | **大批量数据** |

## 📦 核心API

### 1. `excel_to_json_direct` 
**直接将Excel转JSON**

```python
from excel_to_nl import excel_to_json_direct

json_data = excel_to_json_direct(
    filepath="C:\\path\\to\\file.xlsx",
    sheet_name="3-报销",
    serial=1
)
```

### 2. `process_excel_to_mcp_direct`
**单个序号：Excel→JSON→MCP**

```python
from workflow_core import process_excel_to_mcp_direct

mcp_prompt = process_excel_to_mcp_direct(
    excel_path="C:\\path\\to\\file.xlsx",
    sheet_name="3-报销",
    serial=1
)
```

### 3. `batch_process_excel_to_mcp_direct`
**批量处理：Excel所有序号→MCP**

```python
from workflow_core import batch_process_excel_to_mcp_direct

results = batch_process_excel_to_mcp_direct(
    excel_path="C:\\path\\to\\file.xlsx",
    sheet_name="3-报销"
)

for result in results:
    print(f"序号 {result['serial']}")
    print(f"JSON: {result['json_data']}")
    print(f"MCP提示词: {result['mcp_prompt']}")
```

## 🏗️ 系统架构

```
Excel文件 (420财务050823.xlsx)
    │
    ├─ 3-报销 ─────┐
    ├─ 3-差旅      │
    └─ 2-劳务      │
                   ▼
        [excel_to_nl.py]
        excel_to_json_direct()
                   │
                   ▼
              JSON数据
        {
          businessType: "报销业务",
          login: {...},
          project: {...},
          expenses: [...],
          personnel: [...]
        }
                   │
                   ▼
        [workflow_core.py]
        WorkflowCore.build_playwright_prompt_from_data()
                   │
                   ▼
        Playwright MCP提示词
        "1. 打开https://cwcx.uestc.edu.cn/..."
        "2. 在用户名输入框中输入5130008"
        "3. ..."
```

## 📋 支持的工作表类型

### 1. **报销业务** (3-报销)
- ✅ 登录信息 (login)
- ✅ 项目信息 (project)
- ✅ 科目信息 (expenses)
- ✅ 转卡信息 (personnel)
- ✅ 预约信息 (appointment)

### 2. **差旅业务** (3-差旅)
- ✅ 登录信息 (login)
- ✅ 项目信息 (project)
- ✅ 出差人员信息 (travelPerson)
- ✅ 差旅费用信息 (travelExpenses)
- ✅ 转卡信息 (personnel)
- ✅ 预约信息 (appointment)

### 3. **劳务业务** (2-劳务)
- ✅ 登录信息 (login)
- ✅ 项目信息 (project)
- ✅ 劳务信息 (laborInfo)
- ✅ 劳务人员 (laborPerson)
- ✅ 预约信息 (appointment)

## 🔧 Excel列名映射

### 公共字段
- `账号` / `用户名` → `login.username`
- `密码` → `login.password`
- `业务大类` → `businessType`
- `报销项目号` → `project.projectNumber`
- `附件张数` → `project.attachmentCount`
- `支付方式` → `project.paymentMethod`
- `备注` → `project.remarks`
- `特殊事项说明` → `project.special`

### 报销业务特定字段
- `科目` → `expenses[].category`
- `金额` → `expenses[].amount`
- `转卡信息工号` → `personnel[].ID`
- `卡号尾号` → `personnel[].bankCard`
- `个人金额` → `personnel[].amount`

### 差旅业务特定字段
- `出差人` → `travelPerson[].ID`
- `姓名` → `travelPerson[].name`
- `人员类型` → `travelPerson[].personType`
- `省份` → `travelExpenses[].province`
- `起` → `travelExpenses[].startTime`
- `迄` → `travelExpenses[].endTime`
- `飞机票` → `travelExpenses[].airfare`
- `火车票` → `travelExpenses[].trainfare`
- `其他交通费` → `travelExpenses[].otherTransport`
- `住宿费` → `travelExpenses[].accommodation`
- `是否安排伙食` → `travelExpenses[].mealArranged`
- `是否安排交通` → `travelExpenses[].transportArranged`

## 🚀 快速开始

### 测试单个序号

```bash
cd "C:\Users\FH\source\repos\Auto Finan\Auto Finan\LLM_Integration"
python test_direct_excel_to_mcp.py
```

### 在代码中使用

```python
# 方法1：直接获取MCP提示词
from workflow_core import process_excel_to_mcp_direct

prompt = process_excel_to_mcp_direct(
    excel_path="C:\\path\\to\\file.xlsx",
    sheet_name="3-报销",
    serial=1
)
print(prompt)

# 方法2：批量处理
from workflow_core import batch_process_excel_to_mcp_direct

results = batch_process_excel_to_mcp_direct(
    excel_path="C:\\path\\to\\file.xlsx",
    sheet_name="3-报销"
)

for r in results:
    print(f"序号: {r['serial']}")
    print(f"MCP提示词长度: {len(r['mcp_prompt'])} 字符")
```

## 🔄 旧方案保留

如果需要使用LLM生成自然语言描述（可选）：

```python
from excel_to_nl import generate_nl_from_excel_via_llm

nl_summaries = generate_nl_from_excel_via_llm(
    filepath="C:\\path\\to\\file.xlsx",
    sheet_name="3-报销"
)

for i, summary in enumerate(nl_summaries):
    print(f"{i+1}. {summary}")
```

## ⚠️ 注意事项

1. **Excel格式要求**：第一行必须是列名表头，第一列必须是"序号"
2. **数据分组**：同一序号的多行数据会自动合并（如多个科目、多个转卡人）
3. **字段映射**：依赖 `field_type_mapping.xlsx` 文件进行字段类型识别
4. **工作表命名**：自动根据工作表名或"业务大类"列识别业务类型

## 📝 更新日志

### v2.0 (2025-10-09)
- ✅ 新增：直接Excel→JSON转换（跳过LLM）
- ✅ 新增：批量处理API
- ✅ 优化：支持差旅业务和劳务业务
- ✅ 修复：类型注解兼容旧版Python
- ✅ 测试：完整流程验证通过

### v1.0
- 初始版本：Excel→NL→JSON→MCP流程

## 📞 联系方式

如有问题或建议，请联系开发团队。




