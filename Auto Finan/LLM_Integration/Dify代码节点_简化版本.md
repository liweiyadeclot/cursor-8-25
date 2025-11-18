# Dify 代码节点 - 简化内联版本

## ⚠️ 问题

Dify 运行在服务器上，无法访问本地的 `workflow_core.py` 文件，导致 `ModuleNotFoundError`。

## ✅ 解决方案

将所有必要的代码内联到 Dify 代码节点中，不依赖外部文件。

---

## 📝 完整代码（直接复制使用）

### 节点：Excel → MCP 提示词

**直接复制以下代码到 Dify 代码节点中**：

```python
import os
import json
import re
from collections import defaultdict

# 导入 openpyxl（Dify 应该已安装）
try:
    from openpyxl import load_workbook
except ImportError:
    output = {
        "success": False,
        "error": "缺少 openpyxl 模块，请确保 Dify 已安装: pip install openpyxl"
    }

# ============================================================================
# Excel 读取和 JSON 转换函数
# ============================================================================

def excel_to_json_direct_inline(excel_path, sheet_name, serial):
    """从 Excel 读取数据并转换为 JSON"""
    try:
        wb = load_workbook(excel_path, data_only=True)
        ws = wb[sheet_name] if sheet_name else wb.active
        
        # 读取表头
        headers = {}
        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
        for idx, val in enumerate(header_row):
            key = str(val or "").strip()
            if key:
                headers[key] = idx
        
        # 按序号分组
        groups = defaultdict(list)
        current_serial = None
        
        for row in ws.iter_rows(min_row=2, values_only=True):
            serial_idx = headers.get("序号")
            if serial_idx is not None:
                serial_val = row[serial_idx]
                if serial_val is not None:
                    current_serial = str(serial_val).strip()
            
            if current_serial:
                groups[current_serial].append(row)
        
        # 获取指定序号的数据
        target_serial = str(serial).strip()
        if target_serial not in groups:
            return None
        
        rows = groups[target_serial]
        if not rows:
            return None
        
        # 转换为 JSON 结构
        json_data = {}
        
        # 登录信息
        if rows:
            first_row = rows[0]
            login_idx = headers.get("账号") or headers.get("登录界面工号")
            password_idx = headers.get("密码") or headers.get("登录界面密码")
            
            if login_idx is not None:
                json_data["login"] = {
                    "username": str(first_row[login_idx] or "").strip(),
                    "password": str(first_row[password_idx] if password_idx is not None else "").strip()
                }
        
        # 业务大类
        business_type_idx = headers.get("业务大类") or headers.get("选择业务大类")
        if business_type_idx is not None and rows:
            business_type = str(rows[0][business_type_idx] or "").strip()
            if business_type:
                json_data["businessType"] = business_type
        
        # 项目信息
        project_number_idx = headers.get("项目号") or headers.get("报销项目号")
        if project_number_idx is not None and rows:
            project_number = str(rows[0][project_number_idx] or "").strip()
            if project_number:
                json_data["project"] = {
                    "projectNumber": project_number
                }
        
        # 附件张数
        attachment_idx = headers.get("附件张数")
        if attachment_idx is not None and rows:
            attachment = rows[0][attachment_idx]
            if attachment is not None:
                json_data["project"] = json_data.get("project", {})
                json_data["project"]["attachmentCount"] = str(attachment).strip()
        
        # 支付方式
        payment_idx = headers.get("支付方式")
        if payment_idx is not None and rows:
            payment = str(rows[0][payment_idx] or "").strip()
            if payment:
                json_data["project"] = json_data.get("project", {})
                json_data["project"]["paymentMethod"] = payment
        
        # 费用信息（报销业务）
        if json_data.get("businessType") == "报销业务":
            expenses = []
            for row in rows:
                category_idx = headers.get("费用类别") or headers.get("科目")
                amount_idx = headers.get("金额") or headers.get("费用金额")
                
                if category_idx is not None and amount_idx is not None:
                    category = str(row[category_idx] or "").strip()
                    amount = row[amount_idx]
                    
                    if category and amount is not None:
                        expenses.append({
                            "category": category,
                            "amount": str(amount).strip()
                        })
            
            if expenses:
                json_data["expenses"] = expenses
        
        # 人员信息
        personnel = []
        for row in rows:
            person_id_idx = headers.get("学工号") or headers.get("工号")
            card_idx = headers.get("银行卡号尾号") or headers.get("卡号尾号")
            amount_idx = headers.get("金额")
            
            if person_id_idx is not None:
                person_id = str(row[person_id_idx] or "").strip()
                card = str(row[card_idx] if card_idx is not None and row[card_idx] else "").strip()
                amount = str(row[amount_idx] if amount_idx is not None and row[amount_idx] else "").strip()
                
                if person_id:
                    personnel.append({
                        "personId": person_id,
                        "cardLastFour": card,
                        "amount": amount
                    })
        
        if personnel:
            json_data["personnel"] = personnel
        
        # 预约日期
        appointment_idx = headers.get("预约日期") or headers.get("日期")
        if appointment_idx is not None and rows:
            appointment = rows[0][appointment_idx]
            if appointment is not None:
                json_data["appointment"] = {
                    "date": str(appointment).strip()
                }
        
        return json_data if json_data else None
        
    except Exception as e:
        return None

# ============================================================================
# MCP 提示词生成函数
# ============================================================================

def build_playwright_prompt_inline(json_data):
    """从 JSON 数据生成 Playwright MCP 提示词"""
    segments = []
    
    # 添加开头
    segments.append("请你调用Playwright MCP，执行以下命令，一次性执行完")
    segments.append("打开https://cwcx.uestc.edu.cn/WFManager/login.jsp")
    
    # 业务大类
    business_type = json_data.get("businessType", "")
    if business_type:
        segments.append(f"业务大类：{business_type}。以下是需要执行的页面操作：")
    
    # 登录信息
    login = json_data.get("login", {})
    if login.get("username"):
        segments.append(f"在用户名输入框中输入{login['username']}")
    if login.get("password"):
        segments.append(f"在密码输入框中输入{login['password']}")
    
    # 验证码处理
    segments.append("将验证码图片保存至C:\\Users\\FH\\source\\repos\\Auto Finan\\Auto Finan\\LLM_Integration目录下，命名为example.jpg")
    segments.append("运行C:\\Users\\FH\\source\\repos\\Auto Finan\\Auto Finan\\LLM_Integration目录下的OCR.py，得到读取的验证码")
    segments.append("输入验证码")
    segments.append("点击登录按钮")
    segments.append("点击网上预约报账按钮")
    segments.append("点击申请报销单按钮")
    segments.append("点击已阅读并同意按钮")
    
    # 项目信息
    project = json_data.get("project", {})
    if project.get("projectNumber"):
        segments.append(f"在报销项目号输入框中输入{project['projectNumber']}")
    if project.get("attachmentCount"):
        segments.append(f"在附件张数输入框中输入{project['attachmentCount']}")
    if project.get("paymentMethod"):
        segments.append(f"在支付方式下拉框中选择值为\"{project['paymentMethod']}\"")
    segments.append("点击下一步按钮")
    
    # 费用信息（报销业务）
    if business_type == "报销业务":
        expenses = json_data.get("expenses", [])
        for expense in expenses:
            category = expense.get("category", "")
            amount = expense.get("amount", "")
            if category and amount:
                segments.append(f"向{category}输入框填写{amount}")
        segments.append("点击下一步按钮")
    
    # 人员信息
    personnel = json_data.get("personnel", [])
    for i, person in enumerate(personnel):
        person_id = person.get("personId", "")
        card = person.get("cardLastFour", "")
        amount = person.get("amount", "")
        
        if person_id:
            segments.append(f"在学工号输入框中输入{person_id}")
        if card:
            segments.append(f"银行卡号尾号内容为{card}")
        if amount:
            segments.append(f"在金额输入框中输入{amount}")
        
        if i < len(personnel) - 1:
            segments.append("点击提交按钮")
            segments.append("等待页面响应")
        else:
            segments.append("点击下一步按钮")
    
    # 预约信息
    appointment = json_data.get("appointment", {})
    if appointment.get("date"):
        segments.append(f"选择日期预约日期为{appointment['date']}")
    segments.append("点击预约按钮")
    segments.append("点击打印确认单按钮")
    segments.append("调用test_mouse_keyboard.py，执行一个python自动点击的脚本，脚本的第一个参数为保存路径，第二个参数为保存文件名，请你以当前页面中的信息，以报销单号-项目号-金额的格式，输入第二个参数")
    segments.append("等待刚刚运行的脚本运行完毕")
    segments.append("点击返回按钮")
    segments.append("重命名当前读取的提示词文件，将未预约改成已预约")
    
    # 添加序号
    numbered_segments = []
    for i, segment in enumerate(segments, 1):
        if segment.strip():
            numbered_segments.append(f"{i}. {segment.strip()}")
    
    return "\n".join(numbered_segments)

# ============================================================================
# 主处理逻辑
# ============================================================================

# 获取输入变量
excel_path = inputs.get('excel_path', '')
sheet_name = inputs.get('sheet_name', '')
serial = inputs.get('serial', '')

# 验证参数
if not excel_path or not sheet_name or not serial:
    output = {
        "success": False,
        "error": "缺少必要参数"
    }
elif not os.path.exists(excel_path):
    output = {
        "success": False,
        "error": f"Excel 文件不存在: {excel_path}"
    }
else:
    try:
        # 从 Excel 读取并转换为 JSON
        json_data = excel_to_json_direct_inline(excel_path, sheet_name, serial)
        
        if not json_data:
            output = {
                "success": False,
                "error": f"未找到序号 {serial} 的数据"
            }
        else:
            # 生成 MCP 提示词
            mcp_prompt = build_playwright_prompt_inline(json_data)
            
            if not mcp_prompt:
                output = {
                    "success": False,
                    "error": "未能生成有效的 MCP 提示词"
                }
            else:
                output = {
                    "success": True,
                    "mcp_prompt": mcp_prompt,
                    "prompt_length": len(mcp_prompt)
                }
    
    except Exception as e:
        output = {
            "success": False,
            "error": f"处理失败: {str(e)}",
            "error_type": type(e).__name__
        }
```

---

## 📋 使用说明

1. **直接复制**：将上面的完整代码复制到 Dify 代码节点中
2. **不需要修改**：代码已经包含了所有必要的函数
3. **只需要 openpyxl**：确保 Dify 已安装 `openpyxl`

---

## ⚠️ 注意事项

1. **文件路径**：确保 Excel 文件路径在 Dify 服务器上可访问
2. **列名匹配**：代码中使用的列名需要与你的 Excel 文件匹配
3. **简化版本**：这是简化版本，可能不支持所有字段类型

---

## 🔧 如果需要支持更多字段

如果发现某些字段没有被处理，可以修改 `excel_to_json_direct_inline` 函数，添加对应的列名映射。

