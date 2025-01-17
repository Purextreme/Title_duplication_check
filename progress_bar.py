from tqdm import tqdm

class ProgressTracker:
    def __init__(self, total, desc="处理进度"):
        """
        初始化进度跟踪器
        :param total: 总任务数
        :param desc: 进度条描述
        """
        self.total = total
        self.desc = desc
        self.pbar = None
        
    def __enter__(self):
        """创建进度条"""
        self.pbar = tqdm(total=self.total, desc=self.desc)
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        """关闭进度条"""
        if self.pbar:
            self.pbar.close()
            
    def update(self, n=1):
        """更新进度"""
        if self.pbar:
            self.pbar.update(n)
            
    def set_description(self, desc):
        """设置描述"""
        if self.pbar:
            self.pbar.set_description(desc) 