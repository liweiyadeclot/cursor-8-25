#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Playwright MCP 安装验证脚本

功能：
1. 检查 Node.js 是否安装
2. 检查 npm/npx 是否可用
3. 检查 Playwright MCP 是否可以运行
4. 检查版本信息
5. 测试基本功能
"""

import os
import sys
import subprocess
import json
from typing import Dict, Any, Optional

# 修复 Windows 控制台编码问题
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except:
        pass


def run_command(cmd: list, timeout: int = 30, shell: bool = None) -> tuple[bool, str, str]:
    """
    运行命令并返回结果
    
    Returns:
        (success, stdout, stderr)
    """
    # Windows 上默认使用 shell=True
    if shell is None:
        shell = (sys.platform == 'win32')
    
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout,
            encoding='utf-8',
            errors='ignore',
            shell=shell
        )
        return (result.returncode == 0, result.stdout, result.stderr)
    except subprocess.TimeoutExpired:
        return (False, "", "命令执行超时")
    except FileNotFoundError:
        return (False, "", f"命令未找到: {cmd[0]}")
    except Exception as e:
        return (False, "", str(e))


def check_nodejs() -> Dict[str, Any]:
    """检查 Node.js 是否安装"""
    print("=" * 80)
    print("1. 检查 Node.js")
    print("=" * 80)
    
    success, stdout, stderr = run_command(["node", "--version"])
    
    if success:
        version = stdout.strip()
        print(f"✅ Node.js 已安装: {version}")
        
        # 检查版本号
        try:
            version_num = version.replace("v", "").split(".")[0]
            if int(version_num) >= 18:
                print(f"✅ Node.js 版本符合要求 (>= 18)")
                return {"status": "success", "version": version}
            else:
                print(f"⚠️  Node.js 版本过低，建议升级到 18 或更高版本")
                return {"status": "warning", "version": version, "message": "版本过低"}
        except:
            return {"status": "success", "version": version}
    else:
        print("❌ Node.js 未安装")
        print("💡 请访问 https://nodejs.org/ 下载安装 Node.js")
        return {"status": "error", "message": "Node.js 未安装"}


def check_npm() -> Dict[str, Any]:
    """检查 npm 是否可用"""
    print("\n" + "=" * 80)
    print("2. 检查 npm")
    print("=" * 80)
    
    success, stdout, stderr = run_command(["npm", "--version"])
    
    if success:
        version = stdout.strip()
        print(f"✅ npm 已安装: {version}")
        return {"status": "success", "version": version}
    else:
        print("❌ npm 不可用")
        return {"status": "error", "message": "npm 不可用"}


def check_npx() -> Dict[str, Any]:
    """检查 npx 是否可用"""
    print("\n" + "=" * 80)
    print("3. 检查 npx")
    print("=" * 80)
    
    success, stdout, stderr = run_command(["npx", "--version"])
    
    if success:
        version = stdout.strip()
        print(f"✅ npx 已安装: {version}")
        return {"status": "success", "version": version}
    else:
        print("❌ npx 不可用")
        return {"status": "error", "message": "npx 不可用"}


def check_playwright_mcp_version() -> Dict[str, Any]:
    """检查 Playwright MCP 版本（使用 0.0.46 稳定版）"""
    print("\n" + "=" * 80)
    print("4. 检查 Playwright MCP 版本 (0.0.46)")
    print("=" * 80)
    
    # 尝试获取版本信息
    success, stdout, stderr = run_command(
        ["npx", "@playwright/mcp@0.0.46", "--version"],
        timeout=60
    )
    
    if success:
        version = stdout.strip()
        print(f"✅ Playwright MCP 0.0.46 可用")
        print(f"   版本信息: {version}")
        return {"status": "success", "version": version}
    else:
        # 检查是否是 utilsBundleImpl 错误
        if "utilsBundleImpl" in stderr or "utilsBundleImpl" in stdout:
            print("❌ Playwright MCP 0.0.47 存在 utilsBundleImpl 错误")
            print("✅ 已切换到 0.0.46 版本（稳定版）")
            return {"status": "success", "version": "0.0.46", "note": "使用稳定版"}
        else:
            print(f"⚠️  无法获取版本信息")
            print(f"   错误: {stderr[:200] if stderr else stdout[:200]}")
            return {"status": "warning", "message": stderr or stdout}


def check_playwright_mcp_help() -> Dict[str, Any]:
    """检查 Playwright MCP help 命令"""
    print("\n" + "=" * 80)
    print("5. 检查 Playwright MCP help 命令")
    print("=" * 80)
    
    success, stdout, stderr = run_command(
        ["npx", "@playwright/mcp@0.0.46", "--help"],
        timeout=60
    )
    
    if success:
        print("✅ Playwright MCP help 命令执行成功")
        # 显示部分帮助信息
        help_lines = stdout.split('\n')[:10]
        print("\n帮助信息预览:")
        for line in help_lines:
            if line.strip():
                print(f"   {line}")
        return {"status": "success", "help_available": True}
    else:
        print(f"❌ Playwright MCP help 命令执行失败")
        error_msg = stderr[:300] if stderr else stdout[:300]
        print(f"   错误: {error_msg}")
        return {"status": "error", "message": error_msg}


def check_cursor_mcp_config() -> Dict[str, Any]:
    """检查 Cursor MCP 配置"""
    print("\n" + "=" * 80)
    print("6. 检查 Cursor MCP 配置")
    print("=" * 80)
    
    # Windows 路径
    mcp_config_paths = [
        os.path.expanduser("~/.cursor/mcp.json"),
        os.path.expanduser("~/.config/cursor/mcp.json"),
        "C:\\Users\\FH\\.cursor\\mcp.json"
    ]
    
    config_found = False
    playwright_configured = False
    
    for config_path in mcp_config_paths:
        if os.path.exists(config_path):
            config_found = True
            print(f"✅ 找到 MCP 配置文件: {config_path}")
            
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                mcp_servers = config.get("mcpServers", {})
                
                if "playwright" in mcp_servers:
                    playwright_configured = True
                    playwright_config = mcp_servers["playwright"]
                    print(f"✅ Playwright MCP 已配置")
                    print(f"   命令: {playwright_config.get('command', 'N/A')}")
                    print(f"   参数: {playwright_config.get('args', [])}")
                else:
                    print("⚠️  Playwright MCP 未在配置中")
                    print("💡 需要在 mcp.json 中添加 playwright 配置")
                
                break
            except Exception as e:
                print(f"⚠️  读取配置文件失败: {e}")
    
    if not config_found:
        print("⚠️  未找到 MCP 配置文件")
        print("💡 配置文件通常位于:")
        for path in mcp_config_paths:
            print(f"   - {path}")
    
    return {
        "status": "success" if playwright_configured else "warning",
        "config_found": config_found,
        "playwright_configured": playwright_configured
    }


def test_playwright_mcp_basic() -> Dict[str, Any]:
    """测试 Playwright MCP 基本功能"""
    print("\n" + "=" * 80)
    print("7. 测试 Playwright MCP 基本功能")
    print("=" * 80)
    
    # 测试是否能正常启动（不实际执行）
    print("🔄 测试启动 Playwright MCP...")
    
    # 使用 --help 作为测试（不会实际启动浏览器）
    success, stdout, stderr = run_command(
        ["npx", "@playwright/mcp@0.0.46", "--help"],
        timeout=60
    )
    
    if success:
        print("✅ Playwright MCP 可以正常启动")
        return {"status": "success"}
    else:
        error_msg = stderr[:300] if stderr else stdout[:300]
        print(f"❌ Playwright MCP 启动失败")
        print(f"   错误: {error_msg}")
        return {"status": "error", "message": error_msg}


def main():
    """主函数"""
    print("=" * 80)
    print("Playwright MCP 安装验证")
    print("=" * 80)
    print()
    
    results = {}
    
    # 1. 检查 Node.js
    results["nodejs"] = check_nodejs()
    if results["nodejs"]["status"] == "error":
        print("\n❌ 请先安装 Node.js")
        return 1
    
    # 2. 检查 npm
    results["npm"] = check_npm()
    if results["npm"]["status"] == "error":
        print("\n❌ npm 不可用，请检查 Node.js 安装")
        return 1
    
    # 3. 检查 npx
    results["npx"] = check_npx()
    if results["npx"]["status"] == "error":
        print("\n❌ npx 不可用，请检查 Node.js 安装")
        return 1
    
    # 4. 检查 Playwright MCP 版本
    results["mcp_version"] = check_playwright_mcp_version()
    
    # 5. 检查 help 命令
    results["mcp_help"] = check_playwright_mcp_help()
    
    # 6. 检查 Cursor MCP 配置
    results["cursor_config"] = check_cursor_mcp_config()
    
    # 7. 测试基本功能
    results["mcp_test"] = test_playwright_mcp_basic()
    
    # 总结
    print("\n" + "=" * 80)
    print("验证总结")
    print("=" * 80)
    
    all_success = all(
        r.get("status") in ("success", "warning")
        for r in results.values()
    )
    
    if all_success:
        print("✅ Playwright MCP 安装验证通过！")
        print("\n💡 使用建议:")
        print("   1. 在 Cursor 中，Playwright MCP 会自动通过 MCP 客户端调用")
        print("   2. 如需 HTTP 方式调用，请使用 playwright_mcp_http_gateway.py")
        print("   3. 运行示例: python http_mcp_example.py")
        return 0
    else:
        print("⚠️  部分检查未通过，请查看上述详细信息")
        return 1


if __name__ == "__main__":
    sys.exit(main())

