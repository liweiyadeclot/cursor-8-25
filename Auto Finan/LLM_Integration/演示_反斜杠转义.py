"""
演示 JSON 中反斜杠转义的问题
"""
import json
import os

print("=" * 60)
print("JSON 反斜杠转义演示")
print("=" * 60)

# 1. 原始 Windows 路径
original_path = r"C:\Users\FH\source\repos\Auto Finan\Auto Finan\420财务050823.xlsx"
print(f"\n1. 原始路径:")
print(f"   {original_path}")

# 2. Python 字符串（转义形式）
python_escaped = "C:\\Users\\FH\\source\\repos\\Auto Finan\\Auto Finan\\420财务050823.xlsx"
print(f"\n2. Python 转义字符串:")
print(f"   {python_escaped}")
print(f"   相等: {original_path == python_escaped}")

# 3. JSON 序列化
json_data = {"excel_path": python_escaped}
json_str = json.dumps(json_data, ensure_ascii=False)
print(f"\n3. JSON 序列化后:")
print(f"   {json_str}")

# 4. JSON 中的反斜杠数量
backslash_count = json_str.count('\\')
print(f"\n4. JSON 字符串中反斜杠数量: {backslash_count}")

# 5. 多层转义示例
double_escaped = json_str.replace('\\', '\\\\')
print(f"\n5. 再次转义后（模拟多层转义）:")
print(f"   {double_escaped}")
print(f"   反斜杠数量: {double_escaped.count('\\')}")

# 6. JSON 解析
parsed = json.loads(json_str)
parsed_path = parsed["excel_path"]
print(f"\n6. JSON 解析后:")
print(f"   {parsed_path}")
print(f"   相等: {original_path == parsed_path}")

# 7. 处理多层转义
if '\\\\' in double_escaped:
    # 模拟服务端的处理
    temp = double_escaped
    while '\\\\' in temp:
        temp = temp.replace('\\\\', '\\')
    # 解析 JSON
    try:
        fixed_parsed = json.loads(temp)
        fixed_path = fixed_parsed["excel_path"]
        print(f"\n7. 处理多层转义后:")
        print(f"   {fixed_path}")
        print(f"   相等: {original_path == fixed_path}")
    except:
        print(f"\n7. 处理多层转义失败（可能需要其他方法）")

# 8. 使用正斜杠（推荐方案）
forward_slash_path = original_path.replace('\\', '/')
print(f"\n8. 使用正斜杠（推荐）:")
print(f"   {forward_slash_path}")
print(f"   文件存在: {os.path.exists(forward_slash_path)}")

# 9. 正斜杠的 JSON
forward_slash_json = json.dumps({"excel_path": forward_slash_path}, ensure_ascii=False)
print(f"\n9. 正斜杠路径的 JSON:")
print(f"   {forward_slash_json}")
print(f"   反斜杠数量: {forward_slash_json.count('\\')}")

print("\n" + "=" * 60)
print("总结:")
print("=" * 60)
print("1. Windows 路径使用反斜杠，在 JSON 中需要转义")
print("2. 多层转义会导致反斜杠数量成倍增加")
print("3. 使用正斜杠可以避免转义问题")
print("4. 服务端需要正确处理转义")

