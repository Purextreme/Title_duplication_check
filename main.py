import pandas as pd
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import yaml
import os
import nltk
import inflect
from collections import Counter
from tqdm import tqdm
from datetime import datetime

# 下载必要的NLTK数据
try:
    nltk.data.find('corpora/stopwords')
except LookupError:
    nltk.download('stopwords')

class TitleChecker:
    def __init__(self):
        self.config = self._load_config()
        self.p = inflect.engine()
        self.stop_words = set(nltk.corpus.stopwords.words('english'))
        self.log_entries = []
        
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
    
    def normalize_word(self, word):
        """标准化单词，处理单复数"""
        # 清理标点符号
        word = ''.join(c for c in word if c.isalnum() or c == '-')  # 保留连字符
        # 尝试将词转换为单数形式
        singular = self.p.singular_noun(word)
        # 如果是复数形式，返回单数形式；否则返回原词
        return singular if singular else word
    
    def check_title(self, title):
        """检查标题中的重复单词"""
        if not isinstance(title, str):
            return False, []
            
        # 分词并转换为小写
        words = title.lower().split()
        
        # 移除停用词和单个字母
        filtered_words = [w for w in words if w not in self.stop_words and len(w) > 1]
        
        # 将所有单词转换为标准形式（单数）并清理标点
        normalized_words = [self.normalize_word(w) for w in filtered_words]
        
        # 统计词频
        word_count = Counter(normalized_words)
        
        # 找出重复超过阈值的单词
        threshold = self.config['check_settings']['duplicate_threshold']
        duplicates = {word: count for word, count in word_count.items() 
                     if count > threshold}
        
        if duplicates:
            # 创建原始单词到标准化形式的映射
            word_mapping = {}
            for original in filtered_words:
                normalized = self.normalize_word(original)
                if normalized in duplicates:
                    if normalized not in word_mapping:
                        word_mapping[normalized] = set()
                    word_mapping[normalized].add(original)
            
            # 更新duplicates以包含原始形式信息
            detailed_duplicates = {}
            for word, count in duplicates.items():
                original_forms = word_mapping.get(word, {word})
                detailed_duplicates[word] = {
                    'count': count,
                    'original_forms': list(original_forms)
                }
            return True, detailed_duplicates
        
        return False, duplicates
    
    def format_duplicate_info(self, duplicates):
        """格式化重复词信息"""
        if not duplicates:
            return ""
        return "; ".join(f"{word}: {details['count']}次" 
                        for word, details in duplicates.items())
    
    def write_log(self, log_path):
        """写入日志文件"""
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write(f"标题检查报告 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("="*50 + "\n\n")
            for entry in self.log_entries:
                f.write(f"行号: {entry['row']}\n")
                f.write(f"标题: {entry['title']}\n")
                f.write("重复单词:\n")
                for word, details in entry['duplicates'].items():
                    f.write(f"  - {word}:\n")
                    f.write(f"    出现次数: {details['count']}次\n")
                    f.write(f"    原始形式: {', '.join(details['original_forms'])}\n")
                f.write("-"*30 + "\n\n")
    
    def process_excel(self, file_path):
        print("开始处理Excel文件...")
        try:
            # 读取数据
            df = pd.read_excel(file_path, engine='openpyxl')
            wb = load_workbook(file_path)
            ws = wb.active
            
            # 定义红色填充
            red_fill = PatternFill(start_color='FFFF0000', 
                                 end_color='FFFF0000', 
                                 fill_type='solid')
            
            # 获取item-name列
            item_name_col = None
            item_name_col_letter = None
            for idx, col in enumerate(df.columns):
                if col.lower() == 'item-name':
                    item_name_col = col
                    item_name_col_letter = chr(65 + idx)
                    break
                    
            if item_name_col is None:
                raise ValueError("未找到'item-name'列")
            
            # 在item-name列前插入新列
            insert_pos = ord(item_name_col_letter) - 64  # 转换列字母为数字
            ws.insert_cols(insert_pos)
            
            # 更新列字母（因为插入了新列）
            new_col_letter = item_name_col_letter  # 新列使用原来item-name的位置
            item_name_col_letter = chr(ord(item_name_col_letter) + 1)  # item-name列向右移动一位
            
            # 设置新列标题
            ws[f"{new_col_letter}1"] = "重复词检测"
            
            # 确定处理的行数
            if self.config['check_settings']['test_mode']:
                max_rows = min(self.config['check_settings']['test_rows'], len(df))
            else:
                max_rows = len(df)
            
            # 使用tqdm创建进度条
            for index in tqdm(range(max_rows), desc="处理进度"):
                item_name = df.iloc[index][item_name_col]
                if pd.isna(item_name):
                    continue
                
                # 检查标题
                has_duplicates, duplicates = self.check_title(str(item_name))
                
                if has_duplicates:
                    # 记录日志
                    self.log_entries.append({
                        'row': index + 2,  # Excel行号从1开始，还要考虑标题行
                        'title': item_name,
                        'duplicates': duplicates
                    })
                    
                    # 标记item-name列的单元格
                    cell = ws[f"{item_name_col_letter}{index + 2}"]
                    cell.fill = red_fill
                    
                    # 在新列中添加重复词信息
                    info_cell = ws[f"{new_col_letter}{index + 2}"]
                    info_cell.value = self.format_duplicate_info(duplicates)
            
            # 调整新列的宽度
            ws.column_dimensions[new_col_letter].width = 30
            
            # 保存文件
            dir_path = os.path.dirname(file_path)
            file_name = os.path.basename(file_path)
            output_path = os.path.join(dir_path, f"marked_{file_name}")
            log_path = os.path.join(dir_path, f"check_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
            
            print(f"\n保存标记文件到: {output_path}")
            wb.save(output_path)
            
            # 写入日志
            if self.log_entries:
                self.write_log(log_path)
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
