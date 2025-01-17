import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
import yaml
import os
import pandas as pd
from datetime import datetime
from text_analyzer import WordProcessor
from excel_handler import ExcelProcessor
from progress_bar import ProgressTracker

class TitleChecker:
    def __init__(self):
        self.config = self._load_config()
        self.word_processor = WordProcessor()
        self.excel_processor = ExcelProcessor(self.word_processor)
        
    def _load_config(self):
        with open('config.yaml', 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
    
    def select_file(self):
        root = tk.Tk()
        root.withdraw()
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx")]
        )
        return file_path if file_path else None
    
    def process_excel(self, file_path):
        print("开始处理Excel文件...")
        try:
            # 读取工作簿
            wb = load_workbook(file_path)
            ws = wb.active  # 获取活动工作表
            
            # 读取工作表数据
            df = pd.read_excel(file_path, engine='openpyxl')
            
            # 处理工作表
            self.excel_processor.process_worksheet(ws, df, self.config)
            
            # 保存文件
            dir_path = os.path.dirname(file_path)
            file_name = os.path.basename(file_path)
            output_path = os.path.join(dir_path, f"marked_{file_name}")
            log_path = os.path.join(dir_path, f"check_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
            
            print(f"\n保存标记文件到: {output_path}")
            wb.save(output_path)
            
            # 写入日志
            if self.excel_processor.log_entries:
                self.excel_processor.write_log(log_path)
                print(f"检查报告已保存到: {log_path}")
            else:
                print("未发现需要标记的内容")
                
            return output_path
            
        except Exception as e:
            print(f"\n处理文件时出错: {str(e)}")
            raise

def main():
    checker = TitleChecker()
    
    file_path = checker.select_file()
    if not file_path:
        print("未选择文件，程序退出")
        return
        
    try:
        checker.process_excel(file_path)
        print("处理完成")
    except Exception as e:
        print(f"处理过程中出现错误: {str(e)}")

if __name__ == "__main__":
    main()
