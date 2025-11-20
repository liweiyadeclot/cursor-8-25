#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
工作流引擎 - 在Cursor中自动化执行MCP任务

功能：
1. 读取JSON/YAML工作流配置文件
2. 按步骤顺序执行
3. 支持条件判断、循环、变量替换
4. 自动调用MCP HTTP网关
5. 处理错误和重试

使用方法：
    python workflow_engine.py workflow.json
"""

import os
import sys
import json
import yaml
import time
import requests
from typing import Dict, List, Any, Optional
from datetime import datetime
from pathlib import Path

# 确保可以导入本地模块
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

# 默认MCP网关地址
DEFAULT_MCP_ENDPOINT = os.environ.get("MCP_HTTP_ENDPOINT", "http://localhost:3030/mcp/execute")


class WorkflowEngine:
    """工作流执行引擎"""
    
    def __init__(self, config_path: str, mcp_endpoint: str = None):
        """
        初始化工作流引擎
        
        Args:
            config_path: 工作流配置文件路径（JSON或YAML）
            mcp_endpoint: MCP HTTP网关地址
        """
        self.config_path = config_path
        self.mcp_endpoint = mcp_endpoint or DEFAULT_MCP_ENDPOINT
        self.config = self._load_config()
        self.variables = {}  # 工作流变量
        self.logs = []  # 执行日志
        
    def _load_config(self) -> Dict[str, Any]:
        """加载配置文件"""
        path = Path(self.config_path)
        if not path.exists():
            raise FileNotFoundError(f"配置文件不存在: {self.config_path}")
        
        with open(path, 'r', encoding='utf-8') as f:
            if path.suffix.lower() in ['.yaml', '.yml']:
                return yaml.safe_load(f)
            else:
                return json.load(f)
    
    def log(self, message: str, level: str = "INFO"):
        """记录日志"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] [{level}] {message}"
        self.logs.append(log_entry)
        print(log_entry)
    
    def set_variable(self, name: str, value: Any):
        """设置工作流变量"""
        self.variables[name] = value
        self.log(f"设置变量: {name} = {value}")
    
    def get_variable(self, name: str, default: Any = None) -> Any:
        """获取工作流变量"""
        return self.variables.get(name, default)
    
    def resolve_variables(self, text: str) -> str:
        """解析变量占位符 ${variable_name}"""
        if not isinstance(text, str):
            return text
        
        result = text
        for var_name, var_value in self.variables.items():
            placeholder = f"${{{var_name}}}"
            if placeholder in result:
                result = result.replace(placeholder, str(var_value))
        return result
    
    def call_mcp(self, prompt: str, timeout: int = 300, 
                 browser: str = "chrome", headless: bool = False,
                 session_id: Optional[str] = None) -> Dict[str, Any]:
        """
        调用MCP HTTP网关执行Playwright命令
        
        Args:
            prompt: MCP提示词
            timeout: 超时时间（秒）
            browser: 浏览器类型
            headless: 是否无头模式
            session_id: 会话ID（用于保持浏览器会话）
        
        Returns:
            MCP执行结果
        """
        # 解析变量
        prompt = self.resolve_variables(prompt)
        
        self.log(f"调用MCP: {prompt[:100]}...")
        
        payload = {
            "prompt": prompt,
            "timeout": timeout,
            "browser": browser,
            "headless": headless,
        }
        
        if session_id:
            payload["session_id"] = session_id
        
        try:
            response = requests.post(
                self.mcp_endpoint,
                json=payload,
                timeout=timeout + 60  # 额外60秒缓冲
            )
            response.raise_for_status()
            result = response.json()
            
            # 保存session_id到变量
            if "session_id" in result:
                self.set_variable("last_session_id", result["session_id"])
            
            # 保存执行结果到变量
            if "execution_id" in result:
                self.set_variable("last_execution_id", result["execution_id"])
            
            status = result.get("status", "unknown")
            self.log(f"MCP执行完成: {status}")
            
            if result.get("logs"):
                for log_entry in result["logs"][-10:]:  # 只显示最后10条日志
                    self.log(f"  {log_entry}", "MCP")
            
            return result
            
        except requests.exceptions.RequestException as e:
            self.log(f"MCP调用失败: {e}", "ERROR")
            return {
                "status": "error",
                "message": str(e),
                "error_type": type(e).__name__
            }
    
    def execute_step(self, step: Dict[str, Any]) -> bool:
        """
        执行单个工作流步骤
        
        Returns:
            是否成功
        """
        step_type = step.get("type")
        step_name = step.get("name", step_type)
        
        self.log(f"执行步骤: {step_name}")
        
        try:
            if step_type == "mcp":
                # MCP调用步骤
                prompt = step.get("prompt", "")
                timeout = step.get("timeout", 300)
                browser = step.get("browser", "chrome")
                headless = step.get("headless", False)
                session_id = step.get("session_id") or self.get_variable("last_session_id")
                
                result = self.call_mcp(prompt, timeout, browser, headless, session_id)
                
                # 保存结果到变量
                if step.get("save_result_to"):
                    var_name = step["save_result_to"]
                    self.set_variable(var_name, result)
                
                # 检查是否成功
                if result.get("status") in ["success", "partial"]:
                    return True
                else:
                    self.log(f"步骤失败: {result.get('message', '未知错误')}", "ERROR")
                    return False
            
            elif step_type == "set_variable":
                # 设置变量步骤
                var_name = step.get("name")
                var_value = step.get("value")
                # 支持从其他变量获取值
                if isinstance(var_value, str) and var_value.startswith("${") and var_value.endswith("}"):
                    source_var = var_value[2:-1]
                    var_value = self.get_variable(source_var)
                self.set_variable(var_name, var_value)
                return True
            
            elif step_type == "wait":
                # 等待步骤
                seconds = step.get("seconds", 1)
                self.log(f"等待 {seconds} 秒...")
                time.sleep(seconds)
                return True
            
            elif step_type == "condition":
                # 条件判断步骤
                condition = step.get("condition")
                if_true = step.get("if_true", [])
                if_false = step.get("if_false", [])
                
                # 简单的条件判断（支持变量比较）
                condition_met = self._evaluate_condition(condition)
                
                if condition_met:
                    self.log("条件满足，执行 if_true 分支")
                    return self.execute_steps(if_true)
                else:
                    self.log("条件不满足，执行 if_false 分支")
                    return self.execute_steps(if_false)
            
            elif step_type == "loop":
                # 循环步骤
                items = step.get("items", [])
                item_var = step.get("item_var", "item")
                steps = step.get("steps", [])
                
                # 如果 items 是字符串变量引用，从变量中获取
                if isinstance(items, str) and items.startswith("${") and items.endswith("}"):
                    var_name = items[2:-1]
                    items = self.get_variable(var_name, [])
                
                if not isinstance(items, list):
                    self.log(f"循环项必须是列表，当前类型: {type(items)}", "ERROR")
                    return False
                
                for item in items:
                    self.set_variable(item_var, item)
                    self.log(f"循环项: {item}")
                    if not self.execute_steps(steps):
                        # 检查是否允许失败继续
                        if step.get("continue_on_error", False):
                            self.log("循环中某步骤失败但继续", "WARNING")
                            continue
                        else:
                            self.log("循环中某步骤失败，停止循环", "WARNING")
                            return False
                return True
            
            elif step_type == "log":
                # 日志步骤
                message = step.get("message", "")
                message = self.resolve_variables(message)
                self.log(message)
                return True
            
            elif step_type == "script":
                # 执行Python脚本步骤
                script_path = step.get("script")
                if not script_path:
                    self.log("脚本路径未指定", "ERROR")
                    return False
                
                script_path = self.resolve_variables(script_path)
                if not os.path.exists(script_path):
                    self.log(f"脚本文件不存在: {script_path}", "ERROR")
                    return False
                
                # 执行脚本
                import subprocess
                result = subprocess.run(
                    ["python", script_path],
                    capture_output=True,
                    text=True,
                    timeout=step.get("timeout", 60)
                )
                
                if result.returncode == 0:
                    self.log(f"脚本执行成功: {script_path}")
                    if result.stdout:
                        self.log(f"输出: {result.stdout[:200]}")
                    return True
                else:
                    self.log(f"脚本执行失败: {result.stderr[:200]}", "ERROR")
                    return False
            
            elif step_type == "excel_to_prompt":
                # 从Excel生成MCP提示词步骤
                excel_path = step.get("excel_path")
                sheet_name = step.get("sheet_name")
                serial = step.get("serial")
                use_llm = step.get("use_llm", True)
                save_to = step.get("save_to", "mcp_prompt")
                
                if not excel_path:
                    self.log("Excel路径未指定", "ERROR")
                    return False
                
                excel_path = self.resolve_variables(excel_path)
                if sheet_name:
                    sheet_name = self.resolve_variables(sheet_name)
                if serial:
                    serial = self.resolve_variables(serial)
                
                if not os.path.exists(excel_path):
                    self.log(f"Excel文件不存在: {excel_path}", "ERROR")
                    return False
                
                try:
                    # 导入必要的模块
                    import sys
                    current_dir = os.path.dirname(os.path.abspath(__file__))
                    if current_dir not in sys.path:
                        sys.path.insert(0, current_dir)
                    
                    from excel_to_nl import generate_single_nl_from_excel
                    from workflow_core import WorkflowCore
                    
                    self.log(f"从Excel读取数据: {excel_path}, 工作表: {sheet_name}, 序号: {serial}")
                    
                    # 1. 生成自然语言
                    nl_text = generate_single_nl_from_excel(
                        filepath=excel_path,
                        sheet_name=sheet_name,
                        serial=serial,
                        use_llm=use_llm
                    )
                    
                    if not nl_text:
                        self.log(f"序号 {serial} 未找到数据", "ERROR")
                        return False
                    
                    self.log(f"生成的自然语言: {nl_text[:200]}...")
                    
                    # 2. 提取JSON数据
                    workflow = WorkflowCore()
                    json_data = workflow.extract_form_json(nl_text)
                    
                    if not json_data:
                        self.log("无法从自然语言中提取JSON数据", "ERROR")
                        return False
                    
                    self.log(f"提取的JSON数据: {json.dumps(json_data, ensure_ascii=False)[:200]}...")
                    
                    # 3. 生成MCP提示词
                    mcp_prompt = workflow.build_playwright_prompt_from_data(json_data)
                    
                    if not mcp_prompt:
                        self.log("无法生成MCP提示词", "ERROR")
                        return False
                    
                    # 保存到变量
                    self.set_variable(save_to, mcp_prompt)
                    self.set_variable(f"{save_to}_nl", nl_text)
                    self.set_variable(f"{save_to}_json", json_data)
                    
                    self.log(f"✅ MCP提示词已生成并保存到变量: {save_to}")
                    self.log(f"提示词长度: {len(mcp_prompt)} 字符")
                    
                    return True
                    
                except ImportError as e:
                    self.log(f"导入模块失败: {e}", "ERROR")
                    self.log("请确保 excel_to_nl.py 和 workflow_core.py 在同一目录", "ERROR")
                    return False
                except Exception as e:
                    self.log(f"从Excel生成提示词失败: {e}", "ERROR")
                    import traceback
                    self.log(traceback.format_exc(), "ERROR")
                    return False
            
            else:
                self.log(f"未知的步骤类型: {step_type}", "ERROR")
                return False
                
        except Exception as e:
            self.log(f"执行步骤时出错: {e}", "ERROR")
            import traceback
            self.log(traceback.format_exc(), "ERROR")
            return False
    
    def _evaluate_condition(self, condition: str) -> bool:
        """评估条件表达式（简单实现）"""
        # 支持简单的变量比较，如: "${var_name} == 'value'"
        try:
            # 先替换所有变量引用
            import re
            pattern = r'\$\{([^}]+)\}'
            
            # 找到所有变量引用
            matches = re.findall(pattern, condition)
            for var_name in matches:
                var_value = self.get_variable(var_name)
                if var_value is not None:
                    # 根据变量类型决定如何替换
                    if isinstance(var_value, str):
                        # 字符串值：替换为带引号的字符串
                        condition = condition.replace(f"${{{var_name}}}", repr(var_value))
                    elif isinstance(var_value, (int, float)):
                        # 数字值：直接替换
                        condition = condition.replace(f"${{{var_name}}}", str(var_value))
                    elif isinstance(var_value, bool):
                        # 布尔值：替换为 True/False
                        condition = condition.replace(f"${{{var_name}}}", str(var_value))
                    elif isinstance(var_value, dict):
                        # 字典：尝试访问属性，如 ${result.status}
                        # 这里简化处理，只替换整个变量
                        condition = condition.replace(f"${{{var_name}}}", repr(var_value))
                    else:
                        # 其他类型：转换为字符串
                        condition = condition.replace(f"${{{var_name}}}", repr(str(var_value)))
                else:
                    # 变量不存在，替换为空字符串
                    condition = condition.replace(f"${{{var_name}}}", "''")
            
            # 使用安全的eval（限制可用的内置函数）
            safe_dict = {
                "__builtins__": {},
                "True": True,
                "False": False,
                "None": None,
            }
            result = eval(condition, safe_dict)
            return bool(result)
        except Exception as e:
            self.log(f"条件评估失败: {condition}, 错误: {e}", "WARNING")
            # 如果评估失败，返回 False
            return False
    
    def execute_steps(self, steps: List[Dict[str, Any]]) -> bool:
        """执行多个步骤"""
        for step in steps:
            if not self.execute_step(step):
                # 检查是否允许失败继续
                if step.get("continue_on_error", False):
                    self.log("步骤失败但继续执行", "WARNING")
                    continue
                else:
                    return False
        return True
    
    def run(self) -> Dict[str, Any]:
        """运行工作流"""
        workflow_name = self.config.get("name", "未命名工作流")
        self.log(f"开始执行工作流: {workflow_name}")
        
        # 初始化变量
        initial_vars = self.config.get("variables", {})
        for name, value in initial_vars.items():
            self.set_variable(name, value)
        
        # 执行步骤
        steps = self.config.get("steps", [])
        success = self.execute_steps(steps)
        
        # 返回结果
        result = {
            "workflow_name": workflow_name,
            "status": "success" if success else "failed",
            "logs": self.logs,
            "variables": self.variables.copy()
        }
        
        self.log(f"工作流执行完成: {'成功' if success else '失败'}")
        return result


def main():
    """主函数"""
    if len(sys.argv) < 2:
        print("用法: python workflow_engine.py <工作流配置文件>")
        print("示例: python workflow_engine.py workflows/reimburse.json")
        sys.exit(1)
    
    config_path = sys.argv[1]
    mcp_endpoint = sys.argv[2] if len(sys.argv) > 2 else None
    
    try:
        engine = WorkflowEngine(config_path, mcp_endpoint)
        result = engine.run()
        
        # 打印总结
        print("\n" + "="*60)
        print("工作流执行总结")
        print("="*60)
        print(f"状态: {result['status']}")
        print(f"日志条数: {len(result['logs'])}")
        print(f"变量: {result['variables']}")
        
        # 保存结果到文件
        output_file = f"workflow_result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print(f"\n结果已保存到: {output_file}")
        
        sys.exit(0 if result['status'] == 'success' else 1)
        
    except Exception as e:
        print(f"错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()

