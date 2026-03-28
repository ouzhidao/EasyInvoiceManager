#!/usr/bin/env python3
"""
打包脚本 - 将发票整理工具打包成EXE
"""
import os
import sys
import subprocess
import shutil
from pathlib import Path

def clean_build():
    """清理之前的构建文件"""
    dirs_to_remove = ['build', 'dist']
    for dir_name in dirs_to_remove:
        if os.path.exists(dir_name):
            print(f"清理 {dir_name} 目录...")
            shutil.rmtree(dir_name)
    
    # 删除spec文件
    for spec_file in Path('.').glob('*.spec'):
        print(f"删除 {spec_file}...")
        spec_file.unlink()
    
    print("清理完成")

def build_exe():
    """打包EXE"""
    
    # 清理旧文件
    clean_build()
    
    # PyInstaller 参数
    cmd = [
        'pyinstaller',
        '--name=发票整理工具V7.1',  # 程序名称
        '--windowed',  # GUI程序，不显示控制台
        '--onefile',  # 打包成单个EXE文件
        '--clean',  # 清理临时文件
        '--noconfirm',  # 不确认覆盖
        
        # 添加数据文件（如果有图标可以加上）
        # '--icon=app.ico',
        
        # 添加隐藏的导入模块
        '--hidden-import=PyQt5.sip',
        '--hidden-import=PyQt5.QtCore',
        '--hidden-import=PyQt5.QtGui',
        '--hidden-import=PyQt5.QtWidgets',
        '--hidden-import=PyQt5.QtPrintSupport',
        '--hidden-import=PIL',
        '--hidden-import=PIL._imagingtk',
        '--hidden-import=PIL._tkinter_finder',
        '--hidden-import=requests',
        '--hidden-import=openpyxl',
        '--hidden-import=openpyxl.cell._writer',
        '--hidden-import=fitz',  # PyMuPDF
        '--hidden-import=fitz.fitz',
        '--hidden-import=dateutil',
        '--hidden-import=dateutil.tz',
        '--hidden-import=dateutil.parser',
        
        # 添加GUI模块
        '--hidden-import=gui.main_window',
        '--hidden-import=gui.folder_dialog',
        '--hidden-import=gui.print_dialog',
        '--hidden-import=gui.password_dialog',
        '--hidden-import=gui.main_window_v2',
        '--hidden-import=utils.helpers',
        
        # 入口文件
        'main_v6.py'
    ]
    
    print("开始打包...")
    print(f"命令: {' '.join(cmd)}")
    print()
    
    try:
        result = subprocess.run(cmd, check=True, capture_output=False, text=True)
        print("\n✅ 打包成功！")
        return True
    except subprocess.CalledProcessError as e:
        print(f"\n❌ 打包失败: {e}")
        return False

def copy_additional_files():
    """复制额外需要的文件到dist目录"""
    dist_dir = Path('dist')
    
    if not dist_dir.exists():
        print("dist目录不存在")
        return
    
    # 复制说明文档
    files_to_copy = [
        'README_V7.1.txt',
        'requirements.txt',
    ]
    
    for file_name in files_to_copy:
        if Path(file_name).exists():
            shutil.copy(file_name, dist_dir / file_name)
            print(f"复制 {file_name} 到 dist")
    
    # 创建启动说明
    readme_content = """发票整理工具 V7.1 使用说明
========================================

1. 直接双击运行：发票整理工具V7.1.exe

2. 首次使用：
   - 需要先配置百度OCR API密钥
   - 在软件界面点击"配置API"按钮
   - 输入从百度智能云获取的API密钥

3. 基本使用流程：
   - 选择发票来源文件夹
   - 选择输出位置
   - 点击"开始整理"
   - 等待处理完成

4. 注意事项：
   - 需要联网使用（调用百度OCR API）
   - 支持PDF、JPG、PNG等格式
   - 支持ZIP、RAR等压缩包
   - 整理前请确保API配置正确

========================================
版本: V7.1
更新日期: 2026-03-28
"""
    
    with open(dist_dir / '使用说明.txt', 'w', encoding='utf-8') as f:
        f.write(readme_content)
    
    print("创建使用说明文档")

def main():
    """主函数"""
    print("="*60)
    print("发票整理工具 V7.1 打包工具")
    print("="*60)
    print()
    
    # 检查当前目录
    if not Path('main_v6.py').exists():
        print("❌ 错误：当前目录不是项目根目录")
        print("请确保在包含 main_v6.py 的目录中运行此脚本")
        return
    
    # 执行打包
    if build_exe():
        copy_additional_files()
        print()
        print("="*60)
        print("✅ 打包完成！")
        print("="*60)
        print()
        print("输出文件位置: dist/发票整理工具V7.1.exe")
        print()
        print("你可以将 dist 目录中的文件复制给其他同事使用")
    else:
        print()
        print("="*60)
        print("❌ 打包失败，请检查错误信息")
        print("="*60)

if __name__ == '__main__':
    main()
