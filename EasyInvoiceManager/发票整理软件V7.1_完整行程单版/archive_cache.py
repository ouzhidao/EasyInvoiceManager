"""
压缩包处理缓存模块
用于记录已处理的压缩包信息，避免重复扫描
"""
import sqlite3
import hashlib
from pathlib import Path
from datetime import datetime
from typing import Optional, List, Dict


class ArchiveCache:
    """压缩包处理缓存"""
    
    def __init__(self):
        self.db_dir = Path.home() / '.invoice_processor' / 'cache'
        self.db_dir.mkdir(parents=True, exist_ok=True)
        self.db_path = self.db_dir / 'archive_cache.db'
        self._init_database()
    
    def _init_database(self):
        """初始化数据库"""
        conn = sqlite3.connect(str(self.db_path))
        cursor = conn.cursor()
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS archive_cache (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_path TEXT,
                file_md5 TEXT UNIQUE,
                file_size INTEGER,
                file_list TEXT,
                has_invoice INTEGER,
                invoice_count INTEGER,
                process_status TEXT,
                password_required INTEGER,
                processed_time TIMESTAMP,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        cursor.execute('''
            CREATE INDEX IF NOT EXISTS idx_md5 ON archive_cache(file_md5)
        ''')
        
        conn.commit()
        conn.close()
    
    def get_file_md5(self, file_path: str) -> str:
        """计算文件MD5（只读前1MB提高速度）"""
        hash_md5 = hashlib.md5()
        try:
            with open(file_path, "rb") as f:
                # 只读前1MB，大压缩包也快速计算
                hash_md5.update(f.read(1024 * 1024))
        except:
            return ""
        return hash_md5.hexdigest()
    
    def get_cache(self, file_path: str) -> Optional[Dict]:
        """获取缓存信息"""
        file_md5 = self.get_file_md5(file_path)
        
        conn = sqlite3.connect(str(self.db_path))
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT * FROM archive_cache 
            WHERE file_md5 = ?
        ''', (file_md5,))
        
        row = cursor.fetchone()
        conn.close()
        
        if row:
            return {
                'file_path': row[1],
                'file_md5': row[2],
                'file_size': row[3],
                'file_list': row[4].split('\n') if row[4] else [],
                'has_invoice': bool(row[5]),
                'invoice_count': row[6],
                'process_status': row[7],
                'password_required': bool(row[8]),
                'processed_time': row[9]
            }
        return None
    
    def save_cache(self, file_path: str, file_list: List[str], 
                   has_invoice: bool, invoice_count: int,
                   process_status: str = 'pending',
                   password_required: bool = False):
        """保存缓存"""
        file_md5 = self.get_file_md5(file_path)
        file_size = Path(file_path).stat().st_size
        
        conn = sqlite3.connect(str(self.db_path))
        cursor = conn.cursor()
        
        cursor.execute('''
            INSERT OR REPLACE INTO archive_cache 
            (file_path, file_md5, file_size, file_list, has_invoice, 
             invoice_count, process_status, password_required, processed_time)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            file_path, file_md5, file_size, '\n'.join(file_list),
            1 if has_invoice else 0, invoice_count,
            process_status, 1 if password_required else 0,
            datetime.now().isoformat()
        ))
        
        conn.commit()
        conn.close()
    
    def update_status(self, file_path: str, status: str):
        """更新处理状态"""
        file_md5 = self.get_file_md5(file_path)
        
        conn = sqlite3.connect(str(self.db_path))
        cursor = conn.cursor()
        
        cursor.execute('''
            UPDATE archive_cache 
            SET process_status = ?, processed_time = ?
            WHERE file_md5 = ?
        ''', (status, datetime.now().isoformat(), file_md5))
        
        conn.commit()
        conn.close()
    
    def is_processed(self, file_path: str) -> bool:
        """检查是否已处理过"""
        cache = self.get_cache(file_path)
        return cache is not None and cache['process_status'] in ['extracted', 'no_invoice', 'skipped']
    
    def clean_old_cache(self, days: int = 30):
        """清理旧缓存"""
        conn = sqlite3.connect(str(self.db_path))
        cursor = conn.cursor()
        
        cursor.execute('''
            DELETE FROM archive_cache 
            WHERE created_at < datetime('now', '-{} days')
        '''.format(days))
        
        conn.commit()
        conn.close()
