# OCR脚本 - 只输出识别结果的最后一个元素
import ddddocr

# 直接运行OCR
ocr = ddddocr.DdddOcr()
image = open("example.jpg", "rb").read()
result = ocr.classification(image)
print(result)