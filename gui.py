import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import threading
from bartender_worker import BartenderWorker
from config import Config

class BartenderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("模板生成工具")
        self.root.geometry("600x400")
        self.root.resizable(False, False)
        
        self.worker = BartenderWorker()
        self.template_path = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.template_field = tk.StringVar()
        self.printer = tk.StringVar()
        
        self.setup_ui()
        
    def setup_ui(self):
        """设置界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 模板路径
        ttk.Label(main_frame, text="模板路径").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.template_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(main_frame, text="上传模板", command=self.select_template).grid(row=0, column=2)
        
        # Excel路径
        ttk.Label(main_frame, text="Excel路径").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.excel_path, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(main_frame, text="上传Excel", command=self.select_excel).grid(row=1, column=2)
        
        # 输出目录
        ttk.Label(main_frame, text="输出目录:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_dir, width=50).grid(row=2, column=1, padx=5)
        ttk.Button(main_frame, text="选择目录", command=self.select_output_dir).grid(row=2, column=2)
        
        # 模板字段名
        ttk.Label(main_frame, text="模板字段名").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Combobox(main_frame, textvariable=self.template_field, width=47).grid(row=3, column=1, padx=5, sticky=tk.W)
        ttk.Label(main_frame, text=".").grid(row=3, column=2, sticky=tk.W)
        
        # 打印机选择
        ttk.Label(main_frame, text="打印机选择").grid(row=4, column=0, sticky=tk.W, pady=5)
        printer_combo = ttk.Combobox(main_frame, textvariable=self.printer, width=47)
        printer_combo['values'] = Config.DEFAULT_PRINTERS
        printer_combo.grid(row=4, column=1, padx=5, sticky=tk.W)
        ttk.Label(main_frame, text=".").grid(row=4, column=2, sticky=tk.W)
        
        # 分隔线
        ttk.Separator(main_frame, orient='horizontal').grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=20)
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, columnspan=3, pady=10)
        
        self.generate_btn = ttk.Button(button_frame, text="生成文件", command=self.start_generation)
        self.generate_btn.pack(side=tk.LEFT, padx=5)
        
        self.pause_btn = ttk.Button(button_frame, text="暂停生成", command=self.pause_generation, state=tk.DISABLED)
        self.pause_btn.pack(side=tk.LEFT, padx=5)
        
        # 进度条
        self.progress = ttk.Progressbar(main_frame, orient='horizontal', length=500, mode='determinate')
        self.progress.grid(row=7, column=0, columnspan=3, pady=10)
        
        # 进度标签
        self.progress_label = ttk.Label(main_frame, text="0%")
        self.progress_label.grid(row=8, column=0, columnspan=3)
        
        # 状态栏
        self.status_label = ttk.Label(main_frame, text="就绪", foreground="blue")
        self.status_label.grid(row=9, column=0, columnspan=3, pady=10)
        
    def select_template(self):
        """选择模板文件"""
        filename = filedialog.askopenfilename(
            title="选择BTW模板",
            filetypes=[("BarTender文件", "*.btw"), ("所有文件", "*.*")]
        )
        if filename:
            self.template_path.set(filename)
            self.update_status(f"已选择模板: {os.path.basename(filename)}")
            
    def select_excel(self):
        """选择Excel文件"""
        filename = filedialog.askopenfilename(
            title="选择Excel数据文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )
        if filename:
            self.excel_path.set(filename)
            self.update_excel_columns(filename)
            self.update_status(f"已选择数据文件: {os.path.basename(filename)}")
            
    def select_output_dir(self):
        """选择输出目录"""
        directory = filedialog.askdirectory(title="选择输出目录")
        if directory:
            self.output_dir.set(directory)
            self.update_status(f"已选择输出目录: {directory}")
            
    def update_excel_columns(self, excel_path):
        """更新Excel列名到下拉框"""
        try:
            import pandas as pd
            if excel_path.endswith('.csv'):
                df = pd.read_csv(excel_path, nrows=0)
            else:
                df = pd.read_excel(excel_path, nrows=0, engine='openpyxl')
            
            columns = df.colu
