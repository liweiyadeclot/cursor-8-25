#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速运行工作流的便捷脚本

用法:
    python run_workflow.py <工作流文件>
    
示例:
    python run_workflow.py workflows/simple_mcp_test.json
    python run_workflow.py workflows/reimburse_example.json
"""

import sys
import os

# 添加当前目录到路径
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

from workflow_engine import WorkflowEngine

def main():
    if len(sys.argv) < 2:
        print("=" * 60)
        print("工作流执行器")
        print("=" * 60)
        print("\n用法: python run_workflow.py <工作流文件>")
        print("\n示例:")
        print("  python run_workflow.py workflows/simple_mcp_test.json")
        print("  python run_workflow.py workflows/reimburse_example.json")
        print("\n可用工作流:")
        workflows_dir = os.path.join(CURRENT_DIR, "workflows")
        if os.path.exists(workflows_dir):
            for f in os.listdir(workflows_dir):
                if f.endswith(('.json', '.yaml', '.yml')):
                    print(f"  - workflows/{f}")
        sys.exit(1)
    
    workflow_file = sys.argv[1]
    
    # 如果是相对路径，尝试从 workflows 目录查找
    if not os.path.isabs(workflow_file) and not os.path.exists(workflow_file):
        workflows_path = os.path.join(CURRENT_DIR, "workflows", workflow_file)
        if os.path.exists(workflows_path):
            workflow_file = workflows_path
    
    if not os.path.exists(workflow_file):
        print(f"❌ 错误: 工作流文件不存在: {workflow_file}")
        sys.exit(1)
    
    print("=" * 60)
    print(f"执行工作流: {workflow_file}")
    print("=" * 60)
    print()
    
    try:
        engine = WorkflowEngine(workflow_file)
        result = engine.run()
        
        print()
        print("=" * 60)
        print("执行结果")
        print("=" * 60)
        print(f"状态: {'✅ 成功' if result['status'] == 'success' else '❌ 失败'}")
        print(f"日志条数: {len(result['logs'])}")
        if result.get('variables'):
            print(f"\n最终变量:")
            for k, v in result['variables'].items():
                if not k.startswith('last_'):
                    print(f"  {k} = {v}")
        
        sys.exit(0 if result['status'] == 'success' else 1)
        
    except KeyboardInterrupt:
        print("\n\n⚠️  用户中断执行")
        sys.exit(130)
    except Exception as e:
        print(f"\n❌ 错误: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()

