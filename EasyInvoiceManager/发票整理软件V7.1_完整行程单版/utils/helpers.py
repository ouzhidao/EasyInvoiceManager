"""
工具函数集合
"""
import os
import re
import hashlib
from datetime import datetime
from pathlib import Path


def calculate_md5(file_path: str) -> str:
    """计算文件MD5"""
    hash_md5 = hashlib.md5()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_md5.update(chunk)
    return hash_md5.hexdigest()


def sanitize_filename(filename: str) -> str:
    """清理文件名中的非法字符"""
    illegal_chars = r'[\\/:*?"<>|]'
    sanitized = re.sub(illegal_chars, '_', filename)
    if len(sanitized) > 200:
        name, ext = os.path.splitext(sanitized)
        sanitized = name[:200 - len(ext)] + ext
    return sanitized


def extract_short_name(full_name: str, max_length: int = 8) -> str:
    """智能提取开票方简称"""
    if not full_name:
        return "未知"
    
    # 移除常见后缀
    suffixes = ['有限公司', '有限责任公司', '股份有限公司', '公司', '厂', '店', '部', '中心']
    name = full_name
    for suffix in suffixes:
        if name.endswith(suffix):
            name = name[:-len(suffix)]
            break
    
    if len(name) <= max_length:
        return name
    
    return name[:max_length]


def parse_amount(amount_str: str) -> float:
    """解析金额字符串"""
    if not amount_str:
        return 0.0
    
    try:
        return float(amount_str)
    except ValueError:
        pass
    
    # 清理格式
    cleaned = re.sub(r'[¥,\s]', '', amount_str)
    try:
        return float(cleaned)
    except ValueError:
        pass
    
    return 0.0


def standardize_date(date_str: str) -> str:
    """标准化日期格式为YYYYMMDD"""
    if not date_str:
        return ""
    
    digits = re.sub(r'\D', '', date_str)
    if len(digits) == 8:
        return digits
    
    return ""


def get_file_type(file_path: str) -> str:
    """获取文件类型"""
    ext = Path(file_path).suffix.lower()
    if ext in ['.pdf']:
        return 'pdf'
    elif ext in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif']:
        return 'image'
    return 'unknown'


def format_file_size(size_bytes: int) -> str:
    """格式化文件大小"""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.1f} TB"


def ensure_dir(path: str):
    """确保目录存在"""
    Path(path).mkdir(parents=True, exist_ok=True)
