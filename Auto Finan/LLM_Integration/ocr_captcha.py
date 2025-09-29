# OCR脚本 - 识别验证码
import ddddocr
import sys
import os

def recognize_captcha(image_path):
    try:
        # 初始化OCR
        ocr = ddddocr.DdddOcr()
        
        # 读取图片
        if not os.path.exists(image_path):
            print("ERROR: Image file not found")
            return None
            
        with open(image_path, "rb") as f:
            image = f.read()
        
        # 识别验证码
        result = ocr.classification(image)
        
        # 输出结果
        if isinstance(result, str):
            print(result)
            return result
        elif isinstance(result, list) and result:
            print(result[-1])
            return result[-1]
        else:
            print("ERROR: Unable to recognize captcha")
            return None
            
    except Exception as e:
        print(f"ERROR: {str(e)}")
        return None

if __name__ == "__main__":
    if len(sys.argv) > 1:
        image_path = sys.argv[1]
        recognize_captcha(image_path)
    else:
        print("Usage: python ocr_captcha.py <image_path>")
