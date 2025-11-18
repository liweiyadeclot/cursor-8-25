"""
演示 JSON 中反斜杠转义的问题
"""
import json

print("=" * 60)
print("JSON 反斜杠转义演示")
print("=" * 60)

# 1. 原始 Windows 路径
original_path = r"C:\Users\FH\file.xlsx"
print(f"\n1. 原始路径: {original_path}")

# 2. Python 字符串（转义形式）
python_escaped = "C:\\Users\\FH\\file.xlsx"
print(f"2. Python 转义字符串: {python_escaped}")
print(f"   相等: {original_path == python_escaped}")

# 3. JSON 序列化
json_data = {"excel_path": python_escaped}
json_str = json.dumps(json_data)
print(f"\n3. JSON 序列化后:")
print(f"   {json_str}")
print(f"   反斜杠数量: {json_str.count('\\')}")

# 4. 多层转义示例（模拟 Dify 的情况）
double_escaped = json_str.replace('\\', '\\\\')
print(f"\n4. 再次转义后（模拟多层转义）:")
print(f"   {double_escaped}")
print(f"   反斜杠数量: {double_escaped.count('\\')}")

# 5. JSON 解析
parsed = json.loads(json_str)
parsed_path = parsed["excel_path"]
print(f"\n5. JSON 解析后: {parsed_path}")
print(f"   相等: {original_path == parsed_path}")

# 6. 使用正斜杠（推荐方案）
forward_slash_path = original_path.replace('\\', '/')
forward_slash_json = json.dumps({"excel_path": forward_slash_path})
print(f"\n6. 使用正斜杠（推荐）:")
print(f"   路径: {forward_slash_path}")
print(f"   JSON: {forward_slash_json}")
print(f"   反斜杠数量: {forward_slash_json.count('\\')}")

print("\n" + "=" * 60)
print("总结:")
print("=" * 60)
print("1. Windows 路径使用反斜杠，在 JSON 中需要转义")
print("2. 多层转义会导致反斜杠数量成倍增加")
print("3. 使用正斜杠可以避免转义问题")

