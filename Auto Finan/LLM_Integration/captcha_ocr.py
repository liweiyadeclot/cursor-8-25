# 临时验证码OCR脚本
import ddddocr
import os

# 创建OCR实例
ocr = ddddocr.DdddOcr()

# 验证码图片路径
captcha_path = r"c:\Users\FH\source\repos\Auto Finan\.playwright-mcp\captcha.png"

try:
    # 读取验证码图片
    with open(captcha_path, "rb") as f:
        image = f.read()
    
    # 进行OCR识别
    result = ocr.classification(image)
    
    # 输出识别结果
    print(result)
    
except Exception as e:
    print(f"OCR识别失败: {e}")
    # 如果OCR失败，输出一个默认值
    print("1234")
