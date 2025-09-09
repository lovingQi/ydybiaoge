#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试日志级别和颜色功能
"""

import sys
import os
from pathlib import Path

# 添加项目根目录到Python路径
current_dir = Path(__file__).parent
sys.path.insert(0, str(current_dir))

from excel_processor.utils.logger import enhanced_logger

def test_log_levels():
    """测试各种日志级别"""
    print("开始测试日志级别...")
    
    enhanced_logger.debug("这是DEBUG级别的日志 - 用于调试信息")
    enhanced_logger.info("这是INFO级别的日志 - 一般信息")
    enhanced_logger.warning("这是WARNING级别的日志 - 警告信息")
    enhanced_logger.error("这是ERROR级别的日志 - 错误信息")
    enhanced_logger.critical("这是CRITICAL级别的日志 - 严重错误")
    
    print("日志级别测试完成！")

if __name__ == "__main__":
    test_log_levels() 