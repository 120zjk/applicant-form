import logging
from logging.handlers import RotatingFileHandler
import os

def setup_logging(name='app', log_file='application.log'):
    """
    配置並返回一個帶有文件輪替功能的 logger
    """
    # 創建 logger
    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)
    
    # 避免重複添加 handler (Flask 在 debug 模式下可能會重載)
    if not logger.handlers:
        # 定義日誌格式
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        
        # 1. 文件處理器 (File Handler) - 寫入文件
        # maxBytes=1MB, backupCount=5 (保留最近5個備份文件)
        # encoding='utf-8' 確保中文不會亂碼
        file_handler = RotatingFileHandler(
            log_file, 
            maxBytes=1024 * 1024, 
            backupCount=5, 
            encoding='utf-8'
        )
        file_handler.setFormatter(formatter)
        file_handler.setLevel(logging.INFO)
        
        # 2. 控制台處理器 (Stream Handler) - 顯示在螢幕
        stream_handler = logging.StreamHandler()
        stream_handler.setFormatter(formatter)
        stream_handler.setLevel(logging.INFO)
        
        # 添加處理器到 logger
        logger.addHandler(file_handler)
        logger.addHandler(stream_handler)
        
        logger.info(f"Logging setup complete. Log file: {os.path.abspath(log_file)}")
    
    return logger
