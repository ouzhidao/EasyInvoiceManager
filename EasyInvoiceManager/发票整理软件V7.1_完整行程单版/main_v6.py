#!/usr/bin/env python3
"""
发票智能整理工具 - 主入口
"""
import sys
import os

def check_dependencies():
    """检查并安装依赖"""
    required = ['PyQt5', 'openpyxl', 'requests', 'PIL', 'fitz']
    missing = []
    
    for pkg in required:
        try:
            if pkg == 'PIL':
                __import__('PIL')
            elif pkg == 'fitz':
                __import__('fitz')
            else:
                __import__(pkg)
        except ImportError:
            missing.append(pkg)
    
    if missing:
        print(f"缺少依赖: {missing}")
        print("正在安装...")
        import subprocess
        packages = ['PyQt5', 'openpyxl', 'requests', 'pillow', 'pymupdf']
        subprocess.check_call([sys.executable, "-m", "pip", "install"] + packages)
        print("安装完成，请重新运行")
        input("按回车退出...")
        sys.exit(0)

if __name__ == '__main__':
    check_dependencies()
    
    from gui.main_window import main
    main()
