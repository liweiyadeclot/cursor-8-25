#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试预约信息提取
"""

import sys
import os
from datetime import datetime

def test_appointment_extract():
    """测试预约信息提取和文件名生成"""
    print("=== 测试预约信息提取 ===")
    print()
    
    # 模拟提取到的预约信息
    appointment_number = "6824094"
    appointment_time = "2025-09-29 09:00-11:30"
    
    print(f"提取到的预约号: {appointment_number}")
    print(f"提取到的预约时间: {appointment_time}")
    print()
    
    # 清理预约时间格式
    cleaned_time = appointment_time.replace(" ", "").replace(":", "").replace("-", "").replace("&nbsp;", "")
    print(f"清理后的时间: {cleaned_time}")
    print()
    
    # 生成文件名
    file_name = f"{appointment_number}_{cleaned_time}.pdf"
    print(f"生成的文件名: {file_name}")
    print()
    
    # 测试默认值
    print("测试默认值情况:")
    default_number = "null"
    default_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    default_file_name = f"{default_number}_{default_time}.pdf"
    print(f"默认文件名: {default_file_name}")
    print()
    
    print("测试完成！")
    return 0

if __name__ == "__main__":
    sys.exit(test_appointment_extract())
