import pandas as pd
import inflect
from collections import Counter
import re

class WordProcessor:
    def __init__(self, stopwords_file='stopwords.csv'):
        """
        初始化词语处理器
        :param stopwords_file: 停用词文件路径
        """
        self.p = inflect.engine()
        self.stopwords = self._load_stopwords(stopwords_file)
        self.debug_log = []  # 用于存储调试信息
        
    def _load_stopwords(self, file_path):
        """加载停用词表"""
        df = pd.read_csv(file_path)
        stopwords = {}
        for _, row in df.iterrows():
            if row['language'] not in stopwords:
                stopwords[row['language']] = set()
            stopwords[row['language']].add(row['word'])
        return stopwords
    
    def _is_number(self, word):
        """检查是否为数字（包括带逗号的数字）"""
        # 移除逗号
        word = word.replace(',', '')
        # 检查是否为数字或小数
        return bool(re.match(r'^-?\d*\.?\d+$', word))
    
    def normalize_english_word(self, word):
        """标准化英语单词（处理单复数）"""
        # 清理标点符号，保留连字符
        word = ''.join(c for c in word if c.isalnum() or c == '-')
        # 转换为单数形式
        try:
            if self.p.singular_noun(word):  # 如果是复数形式
                return self.p.singular_noun(word)
            return word  # 如果已经是单数形式或不是名词
        except Exception:
            return word  # 如果处理出错，返回原词
    
    def _split_words(self, text):
        """
        拆分单词，处理特殊字符
        """
        # 替换常见分隔符为空格
        text = text.replace(',', ' ').replace(';', ' ').replace(':', ' ')
        
        # 处理特殊字符（保留法语重音符号）
        words = []
        current_word = ''
        for char in text:
            if char.isalnum() or char in '-éèêëàâäôöûüùÉÈÊËÀÂÄÔÖÛÜÙ':
                current_word += char
            else:
                if current_word:
                    words.append(current_word)
                    current_word = ''
                if not char.isspace():  # 如果不是空白字符，作为单独的词
                    words.append(char)
        if current_word:  # 添加最后一个词
            words.append(current_word)
            
        # 过滤空字符串并规范化重音字符
        normalized_words = []
        for word in words:
            word = word.strip()
            if word:
                # 统一重音字符的大小写
                word = ''.join(c.lower() if c in 'ÉÈÊËÀÂÄÔÖÛÜÙéèêëàâäôöûüù' else c for c in word)
                normalized_words.append(word)
                
        return [w for w in normalized_words if w]  # 移除空字符串
    
    def process_english_title(self, title):
        """处理英语标题"""
        if not isinstance(title, str):
            return False, {}
            
        # 分词
        words = self._split_words(title)
        
        # 获取停词集合
        stopwords = self.stopwords.get('english', set())
        
        # 过滤停用词、数字和单个字母
        filtered_words = []
        for word in words:
            word_lower = word.lower()
            if word_lower in stopwords:
                continue
            if self._is_number(word):
                continue
            if len(word) <= 1:
                continue
            filtered_words.append(word)
        
        # 标准化单词（处理单复数）
        normalized_words = [self.normalize_english_word(w) for w in filtered_words]
        
        return self._count_duplicates(normalized_words, filtered_words)
    
    def process_non_english_title(self, title, language):
        """处理非英语标题"""
        if not isinstance(title, str):
            return False, {}
            
        # 分词
        words = self._split_words(title)
        
        # 获取语言对应的停词集合
        stopwords = self.stopwords.get(language.lower(), set())
        if not stopwords and language.lower() != 'english':
            self._log_debug(f"警告：未找到语言 '{language}' 的停词表")
            stopwords = set()
            
        # 调试信息
        self._log_debug(f"\n处理标题: {title}")
        self._log_debug(f"分词结果: {words}")
        self._log_debug(f"停词表 ({language}): {stopwords}")
        
        # 过滤停用词、数字和单个字母
        filtered_words = []
        for word in words:
            word_lower = word.lower()
            if word_lower in stopwords:
                self._log_debug(f"过滤停词: {word}")
                continue
            if self._is_number(word):
                self._log_debug(f"过滤数字: {word}")
                continue
            if len(word) <= 1:
                self._log_debug(f"过滤单字符: {word}")
                continue
            filtered_words.append(word)
            
        self._log_debug(f"过滤后的词: {filtered_words}")
        
        # 对所有语言都使用复数还原
        normalized_words = [self.normalize_english_word(w) for w in filtered_words]
        
        return self._count_duplicates(normalized_words, filtered_words)
    
    def _count_duplicates(self, normalized_words, original_words):
        """统计重复词"""
        # 统计词频
        word_count = Counter(normalized_words)
        
        # 找出重复的词（出现次数大于2）
        duplicates = {word: count for word, count in word_count.items() 
                     if count > 2}
        
        if duplicates:
            # 创建原始单词到标准化形式的映射
            word_mapping = {}
            for original, normalized in zip(original_words, normalized_words):
                if normalized in duplicates:
                    if normalized not in word_mapping:
                        word_mapping[normalized] = set()
                    word_mapping[normalized].add(original)
            
            # 更新duplicates以包含原始形式信息
            detailed_duplicates = {}
            for word, count in duplicates.items():
                original_forms = sorted(list(word_mapping.get(word, {word})))
                detailed_duplicates[word] = {
                    'count': count,
                    'original_forms': original_forms
                }
            return True, detailed_duplicates
        
        return False, {} 
    
    def _log_debug(self, message):
        """
        记录调试信息
        """
        self.debug_log.append(message)
        if len(self.debug_log) > 1000:  # 防止日志太大
            self.debug_log = self.debug_log[-1000:]
    
    def write_debug_log(self, log_file):
        """
        将调试日志写入文件
        """
        if self.debug_log:
            with open(log_file, 'a', encoding='utf-8') as f:
                f.write('\n'.join(self.debug_log) + '\n')
            self.debug_log = []  # 清空日志 