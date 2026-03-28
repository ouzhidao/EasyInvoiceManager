"""
日志记录模块
"""
import logging
from datetime import datetime
from pathlib import Path


class Logger:
    def __init__(self, log_dir: str = None):
        if log_dir is None:
            log_dir = Path.home() / '.invoice_processor' / 'logs'
        
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(parents=True, exist_ok=True)
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        self.log_file = self.log_dir / f'invoice_processor_{timestamp}.log'
        
        self.logger = logging.getLogger('InvoiceProcessor')
        self.logger.setLevel(logging.DEBUG)
        
        file_handler = logging.FileHandler(self.log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(formatter)
        self.logger.addHandler(file_handler)
        
        self.callbacks = []
        self.log_cache = []
    
    def add_callback(self, callback):
        self.callbacks.append(callback)
    
    def _notify(self, level: str, message: str):
        log_entry = {'time': datetime.now().strftime('%H:%M:%S'), 'level': level, 'message': message}
        self.log_cache.append(log_entry)
        
        for callback in self.callbacks:
            try:
                callback(level, message)
            except:
                pass
    
    def info(self, message: str):
        self.logger.info(message)
        self._notify('INFO', message)
    
    def error(self, message: str):
        self.logger.error(message)
        self._notify('ERROR', message)
    
    def get_log_file(self) -> str:
        return str(self.log_file)
