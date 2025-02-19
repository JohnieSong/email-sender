import os
import logging
from datetime import datetime
from pathlib import Path

class Logger:
    _instance = None
    _db = None  # 静态数据库连接
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(Logger, cls).__new__(cls)
            cls._instance._initialize_logger()
        return cls._instance
    
    @classmethod
    def init_db(cls, db):
        """初始化数据库连接"""
        cls._db = db
    
    def _initialize_logger(self):
        """初始化日志配置"""
        # 获取程序运行目录
        current_dir = os.path.dirname(os.path.abspath(__file__))
        self.logs_dir = os.path.join(current_dir, 'logs')
        
        # 创建logs文件夹
        if not os.path.exists(self.logs_dir):
            os.makedirs(self.logs_dir)
            
        # 生成日志文件名
        log_file = os.path.join(self.logs_dir, f"log-{datetime.now().strftime('%Y-%m-%d')}.log")
        
        # 创建logger实例
        self.logger = logging.getLogger('EmailSender')
        self.logger.setLevel(logging.DEBUG)
        
        # 清除已存在的处理器
        if self.logger.handlers:
            self.logger.handlers.clear()
        
        # 创建文件处理器
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        
        # 设置日志格式
        formatter = logging.Formatter(
            '[%(asctime)s] %(levelname)s [%(filename)s:%(lineno)d] - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        file_handler.setFormatter(formatter)
        
        # 添加文件处理器
        self.logger.addHandler(file_handler)
        
        # 只在开发模式下添加控制台处理器
        if os.environ.get('DEVELOPMENT_MODE') == 'true':
            console_handler = logging.StreamHandler()
            console_handler.setLevel(logging.INFO)
            console_handler.setFormatter(formatter)
            self.logger.addHandler(console_handler)
    
    def _log_to_db(self, level, filename, lineno, message):
        """将日志写入数据库"""
        try:
            if self._db:  # 只在数据库连接存在时写入
                self._db.add_system_log(level, filename, lineno, message)
        except Exception as e:
            print(f"写入日志到数据库失败: {str(e)}")

    def debug(self, message):
        self.logger.debug(message)
        self._log_to_db('DEBUG', self.logger.findCaller()[0], self.logger.findCaller()[1], message)
    
    def info(self, message):
        self.logger.info(message)
        self._log_to_db('INFO', self.logger.findCaller()[0], self.logger.findCaller()[1], message)
    
    def warning(self, message):
        self.logger.warning(message)
        self._log_to_db('WARNING', self.logger.findCaller()[0], self.logger.findCaller()[1], message)
    
    def error(self, message):
        self.logger.error(message)
        self._log_to_db('ERROR', self.logger.findCaller()[0], self.logger.findCaller()[1], message)
    
    def critical(self, message):
        self.logger.critical(message)
        self._log_to_db('CRITICAL', self.logger.findCaller()[0], self.logger.findCaller()[1], message) 