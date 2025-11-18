"""
测试路径解码
"""
import json
import os

# 模拟从 Dify 收到的路径（带 Unicode 转义）
test_path = r"C:\\Users\\FH\\source\\repos\\Auto Finan\\Auto Finan\\420\u8d22\u52a1050823.xlsx"

print("原始路径:", repr(test_path))

# 方法 1：使用 json.loads
try:
    decoded_path1 = json.loads(f'"{test_path}"')
    print("方法1 (json.loads):", decoded_path1)
    print("文件存在:", os.path.exists(decoded_path1))
except Exception as e:
    print("方法1 失败:", e)

# 方法 2：手动处理
try:
    import re
    def decode_unicode(match):
        return chr(int(match.group(1), 16))
    decoded_path2 = re.sub(r'\\u([0-9a-fA-F]{4})', decode_unicode, test_path)
    decoded_path2 = decoded_path2.replace('\\\\', '\\')
    decoded_path2 = os.path.normpath(decoded_path2)
    print("方法2 (手动):", decoded_path2)
    print("文件存在:", os.path.exists(decoded_path2))
except Exception as e:
    print("方法2 失败:", e)

# 实际文件路径
actual_path = r"C:\Users\FH\source\repos\Auto Finan\Auto Finan\420财务050823.xlsx"
print("\n实际路径:", actual_path)
print("文件存在:", os.path.exists(actual_path))

