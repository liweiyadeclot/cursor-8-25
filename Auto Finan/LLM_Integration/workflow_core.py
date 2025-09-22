#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
工程核心文件 - 报销自动化工作流程控制
负责管理各个阶段之间的跳转操作和业务逻辑
"""

from typing import Dict, List, Optional, Any, Set
from enum import Enum
import os
import json
import requests

# 可选依赖（首次用到映射表时才需要）
try:
    from openpyxl import Workbook, load_workbook
except Exception:
    Workbook = None  # 延迟报错
    load_workbook = None

class WorkflowStage(Enum):
    """工作流程阶段枚举"""
    LOGIN = "login"
    PROJECT = "project"
    EXPENSE = "expense"
    PERSONNEL = "personnel"
    APPOINTMENT = "appointment"
    TRAVEL = "travel"
    LABOR = "labor"

class WorkflowCore:
    """工作流程核心控制器"""
    
    def __init__(self):
        """初始化工作流程控制器"""
        self._init_stage_transitions()
        self.ollama_base_url: str = os.environ.get("OLLAMA_BASE_URL", "http://localhost:11434")
        self.ollama_model: str = os.environ.get("OLLAMA_MODEL", "qwen2.5:7b")
        self.mapping_excel_path: str = os.path.join(os.path.dirname(__file__), "field_type_mapping.xlsx")
    
    def _init_stage_transitions(self):
        """初始化各阶段跳转操作说明"""
        
        # 1. 登录阶段跳转操作说明
        self.LOGIN_TRANSITION = """
        - 点击登录按钮
        - 点击网上预约报账按钮
        - 点击申请报销单按钮
        - 点击已阅读并同意按钮
        """
        
        # 2. 项目信息阶段跳转操作说明
        self.PROJECT_TRANSITION = """
        - 点击下一步按钮
        """
        
        # 3. 报销科目信息阶段跳转操作说明
        self.EXPENSE_TRANSITION = """
        - 点击下一步按钮
        """
        
        # 4. 报销人员信息阶段跳转操作说明
        self.PERSONNEL_TRANSITION = """
        - 点击下一步按钮
        """
        
        # 5. 预约时间阶段跳转操作说明
        self.APPOINTMENT_TRANSITION = """
        - 点击预约按钮
        - 点击打印确认单按钮
        """
        
        # 6. 差旅信息跳转操作说明
        self.TRAVEL_TRANSITION = """
        - 进入差旅信息填写页面
        - 填写出差人员姓名
        - 选择人员类型
        - 填写出差地点
        - 添加多个出差人员（如有）
        - 验证差旅信息完整性
        - 点击下一步按钮
        - 等待页面跳转
        """
        
        # 7. 劳务信息跳转操作说明
        self.LABOR_TRANSITION = """
        - 进入劳务信息填写页面
        - 选择劳务费类型
        - 填写发放事由
        - 填写劳务金额
        - 添加多个劳务项目（如有）
        - 验证劳务信息完整性
        - 点击提交按钮
        - 等待提交确认
        """
    
    def get_transition_description(self, stage: WorkflowStage) -> str:
        """
        获取指定阶段的跳转操作说明
        
        Args:
            stage: 工作流程阶段
            
        Returns:
            str: 跳转操作说明字符串
        """
        transition_map = {
            WorkflowStage.LOGIN: self.LOGIN_TRANSITION,
            WorkflowStage.PROJECT: self.PROJECT_TRANSITION,
            WorkflowStage.EXPENSE: self.EXPENSE_TRANSITION,
            WorkflowStage.PERSONNEL: self.PERSONNEL_TRANSITION,
            WorkflowStage.APPOINTMENT: self.APPOINTMENT_TRANSITION,
            WorkflowStage.TRAVEL: self.TRAVEL_TRANSITION,
            WorkflowStage.LABOR: self.LABOR_TRANSITION
        }
        
        return transition_map.get(stage, "未知阶段")
    
    def get_all_transitions(self) -> Dict[str, str]:
        """
        获取所有阶段的跳转操作说明
        
        Returns:
            Dict[str, str]: 所有跳转操作说明的字典
        """
        return {
            "login": self.LOGIN_TRANSITION,
            "project": self.PROJECT_TRANSITION,
            "expense": self.EXPENSE_TRANSITION,
            "personnel": self.PERSONNEL_TRANSITION,
            "appointment": self.APPOINTMENT_TRANSITION,
            "travel": self.TRAVEL_TRANSITION,
            "labor": self.LABOR_TRANSITION
        }
    
    def print_all_transitions(self):
        """打印所有阶段的跳转操作说明"""
        print("=== 报销自动化工作流程 - 各阶段跳转操作说明 ===\n")
        
        stages = [
            ("1. 登录阶段", WorkflowStage.LOGIN),
            ("2. 项目信息阶段", WorkflowStage.PROJECT),
            ("3. 报销科目信息阶段", WorkflowStage.EXPENSE),
            ("4. 报销人员信息阶段", WorkflowStage.PERSONNEL),
            ("5. 预约时间阶段", WorkflowStage.APPOINTMENT),
            ("6. 差旅信息阶段", WorkflowStage.TRAVEL),
            ("7. 劳务信息阶段", WorkflowStage.LABOR)
        ]
        
        for stage_name, stage_enum in stages:
            print(f"{stage_name}")
            print(self.get_transition_description(stage_enum))
            print("-" * 80)

    # =========================
    # 新增：信息提取与字段类型映射
    # =========================

    def extract_form_json(self, user_input: str) -> Dict[str, Any]:
        """调用本地Qwen通过Ollama提取结构化JSON。

        Args:
            user_input: 自然语言输入

        Returns:
            提取到的JSON字典（失败时返回带error的字典）
        """
        prompt = self._build_extraction_prompt(user_input)
        try:
            resp = requests.post(
                f"{self.ollama_base_url}/api/generate",
                json={
                    "model": self.ollama_model,
                    "prompt": prompt,
                    "stream": False,
                    "options": {"temperature": 0.1, "top_p": 0.9, "max_tokens": 1200},
                },
                timeout=60,
            )
            if resp.status_code != 200:
                return {"error": f"LLM请求失败: {resp.status_code}", "detail": resp.text}
            data = resp.json()
            text = data.get("response", "").strip()
            # 直接JSON或代码块JSON
            try:
                if text.startswith('{'):
                    return json.loads(text)
            except Exception:
                pass
            # 从```json```代码块提取
            import re
            m = re.search(r"```json\s*(.*?)\s*```", text, re.DOTALL)
            if m:
                return json.loads(m.group(1).strip())
            # 兜底：返回原文
            return {"raw_response": text}
        except Exception as e:
            return {"error": f"LLM请求异常: {e}"}

    def _build_extraction_prompt(self, user_input: str) -> str:
        """与信息提取保持一致的提示词（与测试脚本语义一致，使用英文键名）。"""
        return (
            "你是一个专业的财务报销信息提取助手。请从用户输入的自然语言中提取报销相关信息，"
            "并严格输出JSON且只输出JSON，不要包含任何其他文字。\n\n"
            "数值规范：\n"
            "- 所有金额字段仅输出数字，不要带单位（例如 ‘500’ 而非 ‘500元’）。涉及字段：expenses[].amount、personnel[].amount、labor[].amount。\n"
            "- 附件张数仅输出数字，不要带‘张’（例如 ‘3’ 而非 ‘3张’）。涉及字段：project.attachmentCount。\n\n"
            "键名与结构：\n"
            "{\n"
            "  \"businessType\": \"业务大类\",\n"
            "  \"login\": {\n    \"username\": \"用户名\",\n    \"password\": \"密码\"\n  },\n"
            "  \"project\": {\n    \"projectNumber\": \"项目号\",\n    \"attachmentCount\": \"附件张数(数字)\",\n    \"paymentMethod\": \"支付方式\"\n  },\n"
            "  \"expenses\": [{\n    \"category\": \"科目类型\",\n    \"amount\": \"金额(数字)\"\n  }],\n"
            "  \"personnel\": [{\n    \"name\": \"姓名\",\n    \"ID\": \"学工号\",\n    \"bankCard\": \"银行卡信息\",\n    \"amount\": \"个人金额(数字)\"\n  }],\n"
            "  \"appointment\": {\n    \"date\": \"报销时间\",\n    \"location\": \"地点\"\n  },\n"
            "  \"travel\": [{\n    \"name\": \"姓名\",\n    \"personnelType\": \"人员类型\",\n    \"destination\": \"出差地点\"\n  }],\n"
            "  \"labor\": [{\n    \"laborType\": \"劳务费类型\",\n    \"amount\": \"金额(数字)\",\n    \"reason\": \"发放事由\"\n  }]\n"
            "}\n\n"
            f"用户输入：{user_input}\n"
        )

    def collect_field_paths(self, data: Dict[str, Any]) -> List[str]:
        """从提取到的JSON中收集字段路径（用于生成/校验Excel映射）。

        使用点路径，数组用[]表示，如：expenses[].category
        """
        paths: Set[str] = set()

        def visit(node: Any, prefix: str = ""):
            if isinstance(node, dict):
                for k, v in node.items():
                    new_prefix = f"{prefix}.{k}" if prefix else k
                    visit(v, new_prefix)
            elif isinstance(node, list):
                # 标记为数组
                new_prefix = f"{prefix}[]" if prefix and not prefix.endswith("[]") else prefix
                if node:
                    # 取首元素推断结构
                    visit(node[0], new_prefix)
                else:
                    paths.add(new_prefix)
            else:
                paths.add(prefix)

        visit(data)
        return sorted(paths)

    def ensure_mapping_excel(self, field_paths: List[str]) -> str:
        """确保Excel映射存在；若不存在则创建。若存在则只补充缺失字段。

        Excel格式：第一列=元素名（字段路径），第二列=类型（i/c/r）
        """
        if Workbook is None or load_workbook is None:
            raise RuntimeError("请先安装依赖：pip install openpyxl")

        path = self.mapping_excel_path
        if not os.path.exists(path):
            wb = Workbook()
            ws = wb.active
            ws.title = "mapping"
            ws.append(["element", "type"])  # 表头
            for p in field_paths:
                ws.append([p, "i"])  # 默认输入框，后续可手工改为c/r
            wb.save(path)
            return path

        # 已存在：加载并补齐缺失字段，不覆盖已有配置
        wb = load_workbook(path)
        ws = wb.active
        existing = set()
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and row[0]:
                existing.add(str(row[0]))
        missing = [p for p in field_paths if p not in existing]
        for p in missing:
            ws.append([p, "i"])  # 默认输入框
        if missing:
            wb.save(path)
        return path

    def load_field_type_mapping(self) -> Dict[str, Dict[str, str]]:
        """读取Excel映射为字典：{ 字段路径: { type: i/c/r/d/'' , label: 中文名称 } }

        支持表头：第一列 json字段，第二列 标记，第三列 中文名称
        兼容旧版两列表头：element / type
        """
        if Workbook is None or load_workbook is None:
            raise RuntimeError("请先安装依赖：pip install openpyxl")
        path = self.mapping_excel_path
        if not os.path.exists(path):
            return {}
        wb = load_workbook(path)
        ws = wb.active
        # 解析表头
        headers_raw = [str(c or '').strip() for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
        headers = [h.lower() for h in headers_raw]
        # 头部字段名兼容
        def col_index(names):
            for n in names:
                key = n.lower()
                if key in headers:
                    return headers.index(key)
            return None

        key_idx = col_index(["json字段", "element", "字段", "key", "元素名"]) or 0
        type_idx = col_index(["标记", "type", "类型"]) or 1
        label_idx = col_index(["中文名称", "label", "名称"])  # 可为None

        mapping: Dict[str, Dict[str, str]] = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row:
                continue
            key = str(row[key_idx]).strip() if key_idx is not None and row[key_idx] is not None else ""
            mark = str(row[type_idx]).strip().lower() if type_idx is not None and len(row) > type_idx and row[type_idx] is not None else ""
            label = str(row[label_idx]).strip() if label_idx is not None and len(row) > label_idx and row[label_idx] is not None else ""
            if not key:
                continue
            mapping[key] = {"type": mark, "label": label}
        return mapping

    def build_or_update_mapping_from_input(self, user_input: str) -> Dict[str, str]:
        """一体化流程：调用LLM提取→收集字段→创建/更新Excel→读取映射。

        Returns:
            映射字典 { 字段路径: { type, label } }
        """
        data = self.extract_form_json(user_input)
        if "error" in data:
            return data  # 直接返回错误信息
        # 兼容LLM非JSON返回
        if not isinstance(data, dict):
            return {"error": "LLM未返回JSON"}
        fields = self.collect_field_paths(data)
        self.ensure_mapping_excel(fields)
        return self.load_field_type_mapping()

    # =========================
    # 新增：根据JSON与Excel映射拼接Playwright提示词
    # =========================

    def _flatten_json(self, data: Any, prefix: str = "") -> List[tuple]:
        """展开JSON为 (路径, 值) 列表。数组使用具体索引以保留多项，后续匹配时会归一化为[]。"""
        items: List[tuple] = []
        if isinstance(data, dict):
            for k, v in data.items():
                new_prefix = f"{prefix}.{k}" if prefix else k
                items.extend(self._flatten_json(v, new_prefix))
        elif isinstance(data, list):
            for idx, v in enumerate(data):
                new_prefix = f"{prefix}[{idx}]" if prefix else f"[{idx}]"
                items.extend(self._flatten_json(v, new_prefix))
        else:
            items.append((prefix, data))
        return items

    @staticmethod
    def _normalize_array_path(path: str) -> str:
        """将路径中的具体索引标准化为[]，如 personnel[0].name -> personnel[].name"""
        import re
        return re.sub(r"\[\d+\]", "[]", path)

    def build_playwright_prompt_from_data(self, data: Dict[str, Any]) -> str:
        """根据提取的JSON数据与Excel映射，生成Playwright MCP提示词字符串。"""
        mapping = self.load_field_type_mapping()
        parts: List[str] = []

        # 业务大类作为开头说明（如果有）
        business_type = data.get("businessType")
        if business_type:
            parts.append(f"业务大类：{business_type}。以下是需要执行的页面操作：")

        # 若为“报销业务”，按阶段顺序生成并在每阶段后追加跳转说明
        if (business_type or "").strip() == "报销业务":
            staged = self._build_prompt_reimbursement_flow(data, mapping)
            return staged

        for path, value in self._flatten_json(data):
            if value is None or (isinstance(value, str) and value.strip() == ""):
                continue
            norm_path = self._normalize_array_path(path)
            # 特殊处理：expenses[].amount → 使用同项的category作为控件名称
            if norm_path == "expenses[].amount":
                # 从原始路径提取索引，找到对应category
                import re
                m = re.search(r"expenses\[(\d+)\]\.amount", path)
                if m:
                    idx = int(m.group(1))
                    try:
                        item = (data.get("expenses") or [])[idx]
                        cat = item.get("category") if isinstance(item, dict) else None
                        if cat and str(value).strip() != "":
                            parts.append(f"向{cat}输入框填写{value}")
                            continue  # 已输出定制语句
                    except Exception:
                        pass
            meta = mapping.get(norm_path)
            if not meta:
                continue  # 未登记的字段不参与操作
            mark = (meta.get("type") or "").lower().strip()
            label = meta.get("label") or norm_path  # 无中文名则退回路径
            # 空标记：参考信息，跳过
            if mark == "":
                continue

            if mark == "i":
                parts.append(f"在{label}输入框中输入{value}")
            elif mark == "c":
                parts.append(f"在{label}下拉框中选择{value}")
            elif mark == "r":
                parts.append(f"点击{label}radio button")
            elif mark == "d":
                parts.append(f"选择日期{label}为{value}")
            else:
                # 未知标记，按参考信息忽略
                continue

        return "。".join(parts) + ("。" if parts else "")

    def _build_prompt_reimbursement_flow(self, data: Dict[str, Any], mapping: Dict[str, Dict[str, str]]) -> str:
        """按报销业务的阶段顺序构建提示词，并在每一阶段后拼接跳转说明。"""
        stage_order = [
            ("login", self.LOGIN_TRANSITION),
            ("project", self.PROJECT_TRANSITION),
            ("expenses", self.EXPENSE_TRANSITION),
            ("personnel", self.PERSONNEL_TRANSITION),
            ("appointment", self.APPOINTMENT_TRANSITION),
        ]
        segments: List[str] = []
        # 开头标题
        segments.append("业务大类：报销业务。以下是需要执行的页面操作：")

        for stage_key, transition_text in stage_order:
            actions = self._generate_actions_for_stage(data, stage_key, mapping)
            if actions:
                # 阶段动作句子
                segments.append("。".join(actions) + "。")
            # 阶段跳转说明（无论是否有动作，都附加，便于保持固定流程）
            if transition_text:
                segments.append(transition_text.strip())
        return "\n".join(segments).strip()

    def _generate_actions_for_stage(self, data: Dict[str, Any], stage_key: str, mapping: Dict[str, Dict[str, str]]) -> List[str]:
        """针对某个顶层阶段键（如 'login'、'project'、'expenses'），
        从数据中筛选出该阶段下的所有字段，依据Excel映射生成动作语句。"""
        if stage_key not in data or data.get(stage_key) in (None, ""):
            return []
        stage_actions: List[str] = []

        # 特殊处理：expenses阶段的amount，使用对应category作为控件名
        if stage_key == "expenses" and isinstance(data.get("expenses"), list):
            for item in data.get("expenses") or []:
                if not isinstance(item, dict):
                    continue
                cat = item.get("category")
                amt = item.get("amount")
                if cat is not None and amt not in (None, ""):
                    stage_actions.append(f"向{cat}输入框填写{amt}")
        
        # 遍历所有 (path, value) 对，筛选以 stage_key 开头的路径
        for path, value in self._flatten_json(data):
            if value is None or (isinstance(value, str) and value.strip() == ""):
                continue
            # 只处理该阶段路径
            if not (path == stage_key or path.startswith(stage_key + ".") or path.startswith(stage_key + "[")):
                continue
            # 已经为expenses[].amount生成过定制语句，避免重复
            norm_path = self._normalize_array_path(path)
            if stage_key == "expenses" and norm_path == "expenses[].amount":
                continue

            meta = mapping.get(norm_path)
            if not meta:
                continue
            mark = (meta.get("type") or "").lower().strip()
            label = meta.get("label") or norm_path
            if mark == "":
                continue
            if mark == "i":
                stage_actions.append(f"在{label}输入框中输入{value}")
            elif mark == "c":
                stage_actions.append(f"在{label}下拉框中选择{value}")
            elif mark == "r":
                stage_actions.append(f"点击{label}radio button")
            elif mark == "d":
                stage_actions.append(f"选择日期{label}为{value}")
        return stage_actions

    def build_playwright_prompt_from_input(self, user_input: str) -> Dict[str, Any]:
        """一体化：提取→（可选）更新映射→生成提示词。

        Returns:
            { prompt: str, data: dict }
        """
        data = self.extract_form_json(user_input)
        if not isinstance(data, dict) or "error" in data:
            return {"error": data.get("error", "提取失败"), "data": data}

        # 补齐映射（不会覆盖已有设置）
        fields = self.collect_field_paths(data)
        self.ensure_mapping_excel(fields)

        prompt = self.build_playwright_prompt_from_data(data)
        return {"prompt": prompt, "data": data}

if __name__ == "__main__":
    workflow = WorkflowCore()
    try:
        print("=== 报销自动化 · 交互式生成MCP提示词 ===")
        print("提示：Excel映射表路径为 field_type_mapping.xlsx（与本文件同目录）\n")
        user_text = input("请输入报销相关自然语言（直接回车结束）：\n> ").strip()
        if not user_text:
            print("未输入内容，已退出。")
        else:
            result = workflow.build_playwright_prompt_from_input(user_text)
            if isinstance(result, dict) and result.get("error"):
                print(f"生成失败：{result.get('error')}")
                if result.get("data"):
                    print(result.get("data"))
            else:
                prompt = result.get("prompt", "") if isinstance(result, dict) else ""
                data = result.get("data", {}) if isinstance(result, dict) else {}
                try:
                    print("\n=== 提取到的JSON数据 ===")
                    print(json.dumps(data, ensure_ascii=False, indent=2))
                except Exception:
                    print(data)
                print("\n=== 生成的Playwright MCP提示词 ===")
                print(prompt or "(无可用指令，请检查Excel映射或输入内容)")
    except KeyboardInterrupt:
        print("\n已取消。")
