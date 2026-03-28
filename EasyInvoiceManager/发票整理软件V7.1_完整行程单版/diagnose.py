#!/usr/bin/env python3
"""
诊断脚本 - 检查软件运行环境和性能瓶颈
"""
import os
import sys
import time
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def check_libs():
    """检查必要的库是否安装"""
    print("=" * 50)
    print("检查依赖库")
    print("=" * 50)
    
    libs = {
        'py7zr': '7z压缩包支持',
        'rarfile': 'RAR压缩包支持',
        'win32com': 'Windows快捷方式支持',
        'PyQt5': 'GUI界面',
        'openpyxl': 'Excel生成',
        'PIL': '图像处理',
        'fitz': 'PDF处理 (PyMuPDF)',
    }
    
    for lib, desc in libs.items():
        try:
            __import__(lib)
            print(f"[OK] {lib}: {desc}")
        except ImportError:
            print(f"[MISSING] {lib}: {desc} - 未安装")
    
    print()

def test_archive_recognition():
    """测试压缩包识别"""
    print("=" * 50)
    print("测试压缩包识别")
    print("=" * 50)
    
    from archive_handler import ArchiveHandler
    handler = ArchiveHandler()
    
    test_files = [
        'test.zip',
        'test.7z',
        'test.rar',
        'test.tar.gz',
        'test.tgz',
        'test.tar.bz2',
        'test.pdf',
        'test.jpg',
    ]
    
    print(f"支持的扩展名: {handler.SUPPORTED_EXTS}")
    print(f"has_7z: {handler.has_7z}")
    print(f"has_rar: {handler.has_rar}")
    print()
    
    for f in test_files:
        result = handler.is_archive(f)
        print(f"{f}: {'是压缩包' if result else '不是压缩包'}")
    
    print()

def test_folder_scan_speed(folder_path=None):
    """测试文件夹扫描速度"""
    print("=" * 50)
    print("测试文件夹扫描速度")
    print("=" * 50)
    
    if not folder_path:
        folder_path = str(Path.home())
    
    print(f"测试路径: {folder_path}")
    
    # 测试 Path.iterdir()
    start = time.time()
    try:
        count = 0
        for entry in Path(folder_path).iterdir():
            count += 1
            if count > 10000:
                break
        elapsed = time.time() - start
        print(f"Path.iterdir(): {count} 项, {elapsed:.3f} 秒")
    except Exception as e:
        print(f"Path.iterdir() 失败: {e}")
    
    # 测试 os.scandir()
    start = time.time()
    try:
        count = 0
        for entry in os.scandir(folder_path):
            count += 1
            if count > 10000:
                break
        elapsed = time.time() - start
        print(f"os.scandir(): {count} 项, {elapsed:.3f} 秒")
    except Exception as e:
        print(f"os.scandir() 失败: {e}")
    
    # 测试 os.listdir()
    start = time.time()
    try:
        entries = os.listdir(folder_path)
        elapsed = time.time() - start
        print(f"os.listdir(): {len(entries)} 项, {elapsed:.3f} 秒")
    except Exception as e:
        print(f"os.listdir() 失败: {e}")
    
    # 递归统计文件数
    print("\n递归统计 (这可能需要一些时间)...")
    start = time.time()
    try:
        total_files = 0
        total_dirs = 0
        for root, dirs, files in os.walk(folder_path):
            total_files += len(files)
            total_dirs += len(dirs)
            if total_files > 100000:  # 限制，避免太慢
                break
        elapsed = time.time() - start
        print(f"文件数: {total_files}, 文件夹数: {total_dirs}, 耗时: {elapsed:.3f} 秒")
    except Exception as e:
        print(f"递归统计失败: {e}")
    
    print()

def check_initial_path():
    """检查默认初始路径"""
    print("=" * 50)
    print("检查默认初始路径")
    print("=" * 50)
    
    from config_manager import ConfigManager
    config = ConfigManager()
    
    last_input = config.get_last_path('input')
    last_output = config.get_last_path('output')
    
    print(f"上次输入路径: {last_input or '(无)'}")
    print(f"上次输出路径: {last_output or '(无)'}")
    
    # 检查路径是否存在
    if last_input and os.path.exists(last_input):
        print(f"输入路径存在: 是")
        # 检查文件数量
        try:
            count = len(os.listdir(last_input))
            print(f"输入路径文件数: {count}")
        except:
            print(f"输入路径文件数: 无法读取")
    else:
        print(f"输入路径存在: 否")
    
    print()

def main():
    print("发票整理软件 - 诊断工具")
    print("=" * 50)
    print()
    
    check_libs()
    test_archive_recognition()
    check_initial_path()
    
    # 如果命令行提供了路径，测试该路径
    if len(sys.argv) > 1:
        test_folder_scan_speed(sys.argv[1])
    else:
        # 测试几个常见路径
        test_paths = [
            str(Path.home()),
            'C:\\',
        ]
        for path in test_paths:
            if os.path.exists(path):
                test_folder_scan_speed(path)
                break

if __name__ == '__main__':
    main()
