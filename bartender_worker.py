import os
import sys
import pandas as pd
import win32com.client
import pythoncom
from config import Config

class BartenderWorker:
    def __init__(self):
        self.app = None
        self.format = None
        self.printer_name = None
        self.progress_callback = None
        
    def set_progress_callback(self, callback):
        """设置进度回调函数"""
        self.progress_callback = callback
        
    def update_progress(self, value):
        """更新进度"""
        if self.progress_callback:
            self.progress_callback(value)
            
    def initialize_bartender(self):
        """初始化Bartender应用程序"""
        try:
            pythoncom.CoInitialize()
            self.app = win32com.client.Dispatch("BarTender.Application")
            self.app.Visible = False
            return True
        except Exception as e:
            print(f"初始化Bartender失败: {str(e)}")
            return False
            
    def load_template(self, template_path):
        """加载BTW模板"""
        try:
            self.format = self.app.Formats.Open(template_path)
            return True
        except Exception as e:
            print(f"加载模板失败: {str(e)}")
            return False
            
    def set_printer(self, printer_name):
        """设置打印机"""
        try:
            self.printer_name = printer_name
            # 设置默认打印机
            self.format.PrintSetup.PrinterName = printer_name
            return True
        except Exception as e:
            print(f"设置打印机失败: {str(e)}")
            return False
            
    def read_excel_data(self, excel_path):
        """读取Excel数据"""
        try:
            # 根据文件扩展名选择读取方法
            if excel_path.endswith('.csv'):
                df = pd.read_csv(excel_path)
            else:
                df = pd.read_excel(excel_path, engine='openpyxl')
            
            # 获取列名
            columns = df.columns.tolist()
            data = df.to_dict('records')
            return data, columns
        except Exception as e:
            print(f"读取Excel失败: {str(e)}")
            return None, None
            
    def process_data(self, excel_path, template_field, output_dir):
        """处理数据并生成BTW文件"""
        try:
            # 读取Excel数据
            data, columns = self.read_excel_data(excel_path)
            if not data:
                return False
                
            total_rows = len(data)
            self.update_progress(5)  # 开始处理
            
            # 创建输出目录
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
                
            # 逐行处理数据
            for index, row in enumerate(data):
                # 设置模板中的变量
                for key, value in row.items():
                    try:
                        # 设置命名数据源
                        self.format.SetNamedSubStringValue(key, str(value) if pd.notna(value) else '')
                    except:
                        pass  # 忽略不存在的字段
                
                # 根据模板字段名生成文件名
                if template_field and template_field in row:
                    file_name = f"{row[template_field]}.btw"
                else:
                    file_name = f"label_{index + 1}.btw"
                
                # 保存文件
                output_path = os.path.join(output_dir, file_name)
                self.format.SaveAs(output_path)
                
                # 更新进度
                progress = 5 + (index + 1) * 90 / total_rows
                self.update_progress(int(progress))
            
            self.update_progress(100)
            return True
            
        except Exception as e:
            print(f"处理数据失败: {str(e)}")
            return False
            
    def close(self):
        """关闭Bartender"""
        try:
            if self.format:
                self.format.Close(True)
            if self.app:
                self.app.Quit()
            pythoncom.CoUninitialize()
        except:
            pass
