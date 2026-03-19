# 配置文件
import os

class Config:
    # 默认打印机列表
    DEFAULT_PRINTERS = [
        "TSC TTP-345",
        "TSC TTP-244",
        "Zebra ZT410",
        "Zebra GC420",
        "其他打印机"
    ]
    
    # 支持的Excel格式
    SUPPORTED_EXCEL_FORMATS = ['.xlsx', '.xls', '.csv']
    
    # 支持的模板格式
    SUPPORTED_TEMPLATE_FORMATS = ['.btw']
    
    # 临时文件目录
    TEMP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp')
