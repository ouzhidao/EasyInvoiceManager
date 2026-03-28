"""
查重校验模块
"""
import hashlib


class DuplicateChecker:
    def __init__(self):
        self.index = {}
        self.file_index = {}
    
    def check_duplicate(self, invoice_data: dict, file_md5: str):
        if file_md5 in self.file_index:
            return {'is_duplicate': True, 'index': 0, 'type': 'file'}
        
        key = f"{invoice_data.get('invoice_num_short')}_{invoice_data.get('date')}"
        
        if key not in self.index:
            self.index[key] = 1
            self.file_index[file_md5] = key
            return {'is_duplicate': False, 'index': 0, 'type': 'none'}
        
        self.index[key] += 1
        self.file_index[file_md5] = key
        
        return {
            'is_duplicate': True,
            'index': self.index[key] - 1,
            'type': 'strong'
        }
