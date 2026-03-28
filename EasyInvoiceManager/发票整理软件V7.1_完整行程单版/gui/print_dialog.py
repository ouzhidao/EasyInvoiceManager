"""
打印设置对话框
用于设置各类型发票合并PDF的打印张数，并选择打印机
"""
import os
import subprocess
from pathlib import Path
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
    QSpinBox, QPushButton, QTableWidget, QTableWidgetItem,
    QMessageBox, QHeaderView
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog


class PrintSettingsDialog(QDialog):
    """打印设置对话框"""
    
    def __init__(self, type_summary: dict, output_folder: str, parent=None):
        """
        type_summary: {发票类型: {'count': 数量, 'amount': 金额}}
        output_folder: 输出文件夹路径
        """
        super().__init__(parent)
        self.setWindowTitle("打印设置")
        self.setMinimumSize(600, 400)
        self.setGeometry(200, 200, 700, 500)
        
        self.type_summary = type_summary
        self.output_folder = output_folder
        self.print_settings = {}  # {类型: 打印份数}
        self.merged_pdfs = {}  # {类型: PDF路径}
        
        self.init_ui()
        self.scan_merged_pdfs()
    
    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # 标题
        title_label = QLabel("📄 设置各类型发票合并文档的打印张数")
        title_label.setFont(QFont("Microsoft YaHei", 14, QFont.Bold))
        title_label.setStyleSheet("color: #1976d2; padding: 10px;")
        layout.addWidget(title_label)
        
        # 说明
        info_label = QLabel("系统会自动合并同类型的发票PDF，您可以设置每种类型的打印张数。点击确定后将弹出打印机选择对话框。")
        info_label.setStyleSheet("color: #666; padding: 5px;")
        layout.addWidget(info_label)
        
        # 表格
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels([
            "发票类型", "发票数量", "金额汇总", "合并PDF", "打印张数"
        ])
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.SingleSelection)
        
        # 设置表头
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)
        
        self.table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #ddd;
                font-size: 13px;
            }
            QHeaderView::section {
                background-color: #f5f5f5;
                padding: 10px;
                font-weight: bold;
                font-size: 13px;
                border: 1px solid #ddd;
            }
            QTableWidget::item {
                padding: 8px;
            }
        """)
        
        layout.addWidget(self.table)
        
        # 按钮区
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        # 快速设置按钮
        self.btn_set_all_1 = QPushButton("全部设为 1 张")
        self.btn_set_all_1.setStyleSheet("""
            QPushButton {
                padding: 8px 16px;
                font-size: 12px;
                background-color: #e3f2fd;
                border: 1px solid #1976d2;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #bbdefb;
            }
        """)
        self.btn_set_all_1.clicked.connect(self.set_all_to_1)
        button_layout.addWidget(self.btn_set_all_1)
        
        button_layout.addSpacing(20)
        
        # 取消按钮
        cancel_btn = QPushButton("取消")
        cancel_btn.setStyleSheet("""
            QPushButton {
                padding: 10px 20px;
                font-size: 13px;
                background-color: #f5f5f5;
                border: 1px solid #ccc;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
        """)
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)
        
        # 确定按钮
        ok_btn = QPushButton("✓ 确定并打印")
        ok_btn.setStyleSheet("""
            QPushButton {
                padding: 10px 20px;
                font-size: 13px;
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        ok_btn.clicked.connect(self.on_ok)
        ok_btn.setDefault(True)
        button_layout.addWidget(ok_btn)
        
        layout.addLayout(button_layout)
        
        # 提示
        tip_label = QLabel("💡 提示：点击确定后，将弹出打印机选择对话框，您可以在其中选择打印机、设置打印份数等。")
        tip_label.setStyleSheet("color: #999; font-size: 11px; padding: 5px;")
        layout.addWidget(tip_label)
    
    def scan_merged_pdfs(self):
        """扫描输出文件夹中的合并PDF（支持行程单）"""
        if not self.output_folder or not Path(self.output_folder).exists():
            return
        
        # 填充表格数据
        row = 0
        self.table.setRowCount(len(self.type_summary))
        
        for inv_type, data in sorted(self.type_summary.items()):
            # 查找合并PDF
            merged_pdf = None
            folder_path = Path(self.output_folder)
            
            # 处理行程单文件夹名称
            if inv_type == '行程单':
                type_folder = folder_path / f"行程单_总额{data['amount']:.2f}元"
            else:
                type_folder = folder_path / f"{inv_type}_总额{data['amount']:.2f}元"
            
            if type_folder.exists():
                for pdf_file in type_folder.glob("*_合并.pdf"):
                    merged_pdf = str(pdf_file)
                    self.merged_pdfs[inv_type] = merged_pdf
                    break
            
            # 填充表格
            self.table.setItem(row, 0, QTableWidgetItem(inv_type))
            
            count_item = QTableWidgetItem(f"{data['count']} 张")
            count_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 1, count_item)
            
            amount_item = QTableWidgetItem(f"{data['amount']:.2f} 元")
            amount_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.table.setItem(row, 2, amount_item)
            
            if merged_pdf:
                pdf_item = QTableWidgetItem("✓ 已生成")
                pdf_item.setForeground(Qt.darkGreen)
            else:
                pdf_item = QTableWidgetItem("✗ 未生成")
                pdf_item.setForeground(Qt.red)
            pdf_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 3, pdf_item)
            
            # 打印张数选择
            spin_box = QSpinBox()
            spin_box.setRange(0, 99)
            spin_box.setValue(1 if merged_pdf else 0)
            spin_box.setEnabled(merged_pdf is not None)
            spin_box.setStyleSheet("""
                QSpinBox {
                    padding: 5px;
                    font-size: 13px;
                    border: 1px solid #ddd;
                    border-radius: 3px;
                }
            """)
            spin_box.setMinimumWidth(80)
            self.table.setCellWidget(row, 4, spin_box)
            self.print_settings[inv_type] = spin_box
            
            row += 1
        
        self.table.resizeRowsToContents()
    
    def set_all_to_1(self):
        """全部设为1张"""
        for spin_box in self.print_settings.values():
            if spin_box.isEnabled():
                spin_box.setValue(1)
    
    def on_ok(self):
        """确定并打印 - 弹出打印机选择对话框"""
        # 收集打印设置
        print_list = []
        for inv_type, spin_box in self.print_settings.items():
            count = spin_box.value()
            if count > 0 and inv_type in self.merged_pdfs:
                print_list.append({
                    'type': inv_type,
                    'pdf_path': self.merged_pdfs[inv_type],
                    'count': count
                })
        
        if not print_list:
            QMessageBox.information(self, "提示", "没有需要打印的文档（打印张数都设为0）")
            self.accept()
            return
        
        # 弹出打印机选择对话框
        self.show_print_dialog_and_execute(print_list)
    
    def show_print_dialog_and_execute(self, print_list):
        """显示打印机选择对话框并执行打印"""
        # 创建打印机对象
        printer = QPrinter(QPrinter.HighResolution)
        
        # 创建打印对话框
        print_dialog = QPrintDialog(printer, self)
        print_dialog.setWindowTitle("选择打印机")
        
        # 设置打印选项
        print_dialog.setOption(QPrintDialog.PrintToFile, True)
        print_dialog.setOption(QPrintDialog.PrintPageRange, False)
        print_dialog.setOption(QPrintDialog.PrintCollateCopies, True)
        
        # 显示打印对话框
        if print_dialog.exec_() != QPrintDialog.Accepted:
            # 用户取消了打印
            return
        
        # 用户确认了打印，执行打印操作
        self.execute_print_with_printer(print_list, printer)
    
    def execute_print_with_printer(self, print_list, printer):
        """使用选定的打印机执行打印"""
        success_count = 0
        failed_items = []
        manual_print_items = []  # 需要手动打印的项目
        
        # 获取选定的打印机名称
        printer_name = printer.printerName()
        
        for item in print_list:
            try:
                pdf_path = item['pdf_path']
                count = item['count']
                
                if os.name == 'nt':  # Windows
                    printed = self._print_pdf_windows(pdf_path, printer_name, count)
                    if printed:
                        success_count += 1
                    else:
                        # 静默打印失败，标记为需要手动打印
                        manual_print_items.append(item)
                else:  # macOS/Linux
                    for i in range(count):
                        subprocess.run(['lpr', '-P', printer_name, pdf_path], check=True)
                    success_count += 1
                
            except Exception as e:
                failed_items.append(f"{item['type']}: {str(e)}")
        
        # 关闭设置对话框
        self.accept()
        
        # 显示结果
        self._show_print_result(success_count, failed_items, manual_print_items, printer_name)
    
    def _print_pdf_windows(self, pdf_path: str, printer_name: str, count: int) -> bool:
        """
        Windows下打印PDF，尝试多种方式
        返回: 是否成功使用命令行打印
        """
        # 方法1: 使用SumatraPDF（推荐，轻量且支持命令行指定打印机）
        sumatra_paths = [
            r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
            r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
        ]
        
        for sumatra_path in sumatra_paths:
            if os.path.exists(sumatra_path):
                try:
                    for i in range(count):
                        cmd = [
                            sumatra_path,
                            "-print-to", printer_name,
                            "-print-settings", "fit",
                            pdf_path
                        ]
                        subprocess.run(cmd, check=True, capture_output=True, timeout=60)
                    return True
                except Exception:
                    continue  # 尝试下一个方法
        
        # 方法2: 使用Adobe Acrobat Reader
        adobe_paths = [
            r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
            r"C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
            r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
            r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
        ]
        
        for adobe_path in adobe_paths:
            if os.path.exists(adobe_path):
                try:
                    for i in range(count):
                        # /t 参数: 打印到指定打印机后关闭
                        cmd = [adobe_path, "/t", pdf_path, printer_name]
                        subprocess.run(cmd, check=True, capture_output=True, timeout=60)
                    return True
                except Exception:
                    continue
        
        # 没有找到支持的PDF阅读器，返回False表示需要手动打印
        return False
    
    def _show_print_result(self, success_count: int, failed_items: list, manual_print_items: list, printer_name: str):
        """显示打印结果"""
        # 构建消息
        messages = []
        
        if success_count > 0:
            messages.append(f"✅ 成功打印 {success_count} 个类型到打印机: {printer_name}")
        
        if manual_print_items:
            messages.append(f"\n⚠️ 以下 {len(manual_print_items)} 个类型需要手动打印:")
            for item in manual_print_items:
                messages.append(f"  • {item['type']}: {item['count']} 张")
            messages.append("\n原因: 未检测到支持命令行打印的PDF阅读器（SumatraPDF 或 Adobe Reader）")
            messages.append("系统将为您打开这些PDF文件，请手动选择打印机进行打印。")
        
        if failed_items:
            messages.append(f"\n❌ 打印失败的类型:")
            for item in failed_items:
                messages.append(f"  • {item}")
        
        # 显示消息框
        if failed_items:
            QMessageBox.warning(self, "打印结果", "\n".join(messages))
        elif manual_print_items:
            QMessageBox.information(self, "打印结果", "\n".join(messages))
            # 打开需要手动打印的PDF文件
            for item in manual_print_items:
                try:
                    os.startfile(item['pdf_path'])
                except Exception:
                    pass
        else:
            QMessageBox.information(self, "打印完成", "\n".join(messages))
    
    @staticmethod
    def show_print_dialog(type_summary: dict, output_folder: str, parent=None) -> bool:
        """
        显示打印设置对话框
        返回：用户是否点击了确定
        """
        dialog = PrintSettingsDialog(type_summary, output_folder, parent)
        result = dialog.exec_()
        return result == QDialog.Accepted
