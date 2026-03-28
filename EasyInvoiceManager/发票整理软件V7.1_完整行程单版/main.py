"""
发票智能整理工具 V7.0
支持：增值税发票 + 12306铁路电子客票
OCR：PaddleOCR本地识别 + 百度API备用
"""
import sys
from pathlib import Path

# 确保可以导入本地模块
sys.path.insert(0, str(Path(__file__).parent))

def main():
    try:
        from PyQt5.QtWidgets import QApplication
        from gui.main_window_v2 import MainWindowV2
        
        app = QApplication(sys.argv)
        app.setApplicationName("发票智能整理工具 V7.0")
        app.setApplicationVersion("7.0.0")
        
        window = MainWindowV2()
        window.show()
        
        sys.exit(app.exec_())
        
    except ImportError as e:
        print(f"导入错误: {e}")
        print("\n请确保已安装所有依赖:")
        print("  pip install -r requirements.txt")
        print("\n如果PaddleOCR安装失败，可以尝试:")
        print("  pip install paddleocr -i https://pypi.tuna.tsinghua.edu.cn/simple")
        input("\n按回车键退出...")
        sys.exit(1)
    except Exception as e:
        print(f"启动错误: {e}")
        import traceback
        traceback.print_exc()
        input("\n按回车键退出...")
        sys.exit(1)


if __name__ == '__main__':
    main()
