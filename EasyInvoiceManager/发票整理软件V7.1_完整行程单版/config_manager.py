"""
配置管理模块
"""
import os
import json
import base64
from pathlib import Path


class ConfigManager:
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialized = False
        return cls._instance
    
    def __init__(self):
        if self._initialized:
            return
            
        self._initialized = True
        self.config_dir = Path.home() / '.invoice_processor'
        self.config_file = self.config_dir / 'config.json'
        self.config = {}
        
        self._ensure_config_dir()
        self.load_config()
    
    def _ensure_config_dir(self):
        self.config_dir.mkdir(parents=True, exist_ok=True)
    
    def _encrypt(self, text: str) -> str:
        if not text:
            return ""
        key = b'invoice2025'
        text_bytes = text.encode('utf-8')
        encrypted = bytes([text_bytes[i] ^ key[i % len(key)] for i in range(len(text_bytes))])
        return base64.b64encode(encrypted).decode('utf-8')
    
    def _decrypt(self, encrypted_text: str) -> str:
        if not encrypted_text:
            return ""
        try:
            key = b'invoice2025'
            encrypted = base64.b64decode(encrypted_text.encode('utf-8'))
            decrypted = bytes([encrypted[i] ^ key[i % len(key)] for i in range(len(encrypted))])
            return decrypted.decode('utf-8')
        except:
            return ""
    
    def load_config(self):
        default_config = {
            'paths': {
                'last_input_path': '',
                'last_output_path': str(Path.home() / '发票整理'),
            },
            'api': {
                'baidu_app_id': '',
                'baidu_api_key': '',
                'baidu_secret_key': '',
                'use_paddle_backup': True
            },
            'processing': {
                'confidence_threshold': 85,
                'max_retry': 3,
            }
        }
        
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded = json.load(f)
                    self.config = {**default_config, **loaded}
                    # 加载时立即解密API密钥，确保内存中始终是明文
                    api = self.config.get('api', {})
                    for key in ['baidu_app_id', 'baidu_api_key', 'baidu_secret_key']:
                        if key in api and api[key]:
                            api[key] = self._decrypt(api[key])
            except:
                self.config = default_config
        else:
            self.config = default_config
    
    def save_config(self):
        try:
            config_to_save = self.config.copy()
            api = config_to_save.get('api', {})
            for key in ['baidu_app_id', 'baidu_api_key', 'baidu_secret_key']:
                if key in api and api[key]:
                    api[key] = self._encrypt(api[key])
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config_to_save, f, ensure_ascii=False, indent=2)
            return True
        except:
            return False
    
    def get(self, section: str, key: str, default=None):
        return self.config.get(section, {}).get(key, default)
    
    def set(self, section: str, key: str, value):
        if section not in self.config:
            self.config[section] = {}
        self.config[section][key] = value
    
    def get_api_credentials(self):
        api = self.config.get('api', {})
        return {
            'app_id': self._decrypt(api.get('baidu_app_id', '')),
            'api_key': self._decrypt(api.get('baidu_api_key', '')),
            'secret_key': self._decrypt(api.get('baidu_secret_key', ''))
        }
    
    def set_api_credentials(self, app_id: str, api_key: str, secret_key: str):
        self.config['api']['baidu_app_id'] = app_id
        self.config['api']['baidu_api_key'] = api_key
        self.config['api']['baidu_secret_key'] = secret_key
    
    def get_last_path(self, path_type: str = 'input') -> str:
        key = f'last_{path_type}_path'
        return self.get('paths', key, '')
    
    def set_last_path(self, path_type: str, path: str):
        self.set('paths', f'last_{path_type}_path', path)
