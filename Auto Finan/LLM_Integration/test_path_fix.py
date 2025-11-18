"""
测试路径修复
"""
import json
import os

# 模拟从 Dify 收到的路径（带 Unicode 转义和双反斜杠）
test_path = r"C:\\Users\\FH\\source\\repos\\Auto Finan\\Auto Finan\\420\u8d22\u52a1050823.xlsx"

print("=" * 60)
print("原始路径:")
print(f"  repr: {repr(test_path)}")
print(f"  值: {test_path}")

# 步骤 1: JSON 解码 Unicode
try:
    decoded = json.loads(f'"{test_path}"')
    print("\n步骤 1 - JSON 解码后:")
    print(f"  repr: {repr(decoded)}")
    print(f"  值: {decoded}")
except Exception as e:
    print(f"\n步骤 1 失败: {e}")
    decoded = test_path

# 步骤 2: 处理反斜杠
fixed = decoded
while '\\\\' in fixed:
    fixed = fixed.replace('\\\\', '\\')
fixed = fixed.replace('\\/', '/')

print("\n步骤 2 - 处理反斜杠后:")
print(f"  repr: {repr(fixed)}")
print(f"  值: {fixed}")

# 步骤 3: 规范化
normalized = os.path.normpath(fixed)
print("\n步骤 3 - 规范化后:")
print(f"  repr: {repr(normalized)}")
print(f"  值: {normalized}")

# 验证文件
print("\n" + "=" * 60)
print("文件验证:")
print(f"  文件存在: {os.path.exists(normalized)}")

# 实际路径
actual_path = r"C:\Users\FH\source\repos\Auto Finan\Auto Finan\420财务050823.xlsx"
print(f"\n实际路径:")
print(f"  repr: {repr(actual_path)}")
print(f"  值: {actual_path}")
print(f"  文件存在: {os.path.exists(actual_path)}")
print(f"  路径匹配: {normalized == actual_path}")

