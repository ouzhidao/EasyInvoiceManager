"""
主窗口 - 完整版（含统计功能）
"""
import sys
import os
from pathlib import Path
from datetime import datetime

from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QFileDialog,
    QTextEdit, QProgressBar, QMessageBox, QGroupBox,
    QTableWidget, QTableWidgetItem, QApplication,
    QTabWidget, QTreeWidget, QTreeWidgetItem, QSplitter,
    QMenuBar, QAction, QDialog, QFormLayout
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QDate, QUrl
from PyQt5.QtGui import QIcon, QColor
from PyQt5.QtGui import QDesktopServices

sys.path.insert(0, str(Path(__file__).parent.parent))

from config_manager import ConfigManager
from logger import Logger
from ocr_engine import OCREngine
from data_parser import DataParser
from duplicate_checker import DuplicateChecker
from file_organizer import FileOrganizer
from excel_generator import ExcelGenerator
from statistics_manager import StatisticsManager, InvoiceRecord
from utils.helpers import calculate_md5

# 导入自定义对话框
from gui.folder_dialog import FolderDialog


class WorkerThread(QThread):
    progress_updated = pyqtSignal(int, int, str)
    file_done = pyqtSignal(dict)
    finished_ok = pyqtSignal(dict)
    error = pyqtSignal(str)
    
    def __init__(self, input_path, output_path, stats_manager: StatisticsManager, 
                 files_to_process: list = None, temp_extracted_files: list = None):
        super().__init__()
        self.input_path = input_path
        self.output_path = output_path
        self.stats_manager = stats_manager
        self.files_to_process = files_to_process or []
        self.temp_extracted_files = temp_extracted_files or []
        self.running = True
    
    def run(self):
        try:
            ocr = OCREngine()
            parser = DataParser()
            checker = DuplicateChecker()
            organizer = FileOrganizer(self.output_path)
            excel = ExcelGenerator()
            
            organizer.create_folder_structure()
            
            # 使用预处理好的文件列表
            files = self.files_to_process
            
            total = len(files)
            if total == 0:
                self.error.emit("未找到有效的发票文件")
                return
            
            # 开始统计任务
            self.stats_manager.start_task()
            results = []
            
            for i, file_path in enumerate(files, 1):
                if not self.running:
                    break
                
                self.progress_updated.emit(i, total, Path(file_path).name)
                
                # 创建记录 - 每个文件一个独立记录
                record = None
                try:
                    record = InvoiceRecord(
                        task_id=self.stats_manager.current_task_id,
                        file_name=Path(file_path).name,
                        file_path=file_path,
                        invoice_num='',
                        amount=0,
                        invoice_type='',
                        invoice_date='',
                        seller_name='',
                        buyer_name='',  # 添加购买方名称字段
                        is_success=False,
                        is_duplicate=False,
                        error_msg='',
                        process_time=datetime.now().isoformat(),
                        output_path=''
                    )
                except Exception as e:
                    # 如果创建记录失败，记录错误但继续处理
                    print(f"创建记录失败 {file_path}: {e}")
                    result = {
                        'success': False,
                        'file_path': file_path,
                        'error': f"创建记录失败: {str(e)}",
                        'is_not_invoice': False
                    }
                    results.append(result)
                    self.file_done.emit(result)
                    continue
                
                # 处理单个文件 - 独立的异常处理
                process_error = None
                try:
                    ocr_result = ocr.recognize(file_path)
                    data = parser.parse(ocr_result)
                    
                    # 检查是否为有效发票（关键字段缺失太多则判定为非发票）
                    if not data.get('is_valid'):
                        # 检查是否完全不是发票（关键字段都缺失）
                        has_invoice_num = bool(data.get('invoice_num_full'))
                        has_amount = data.get('amount', 0) > 0
                        has_date = bool(data.get('date'))
                        
                        if not has_invoice_num and not has_amount and not has_date:
                            raise Exception("识别出来不是发票")
                        else:
                            raise Exception("识别结果不完整")
                    
                    file_md5 = calculate_md5(file_path)
                    dup_info = checker.check_duplicate(data, file_md5)
                    
                    ext = Path(file_path).suffix
                    new_name = organizer.generate_filename(data, dup_info, ext)
                    
                    # 确保 invoice_type 不为 None
                    invoice_type = data.get('invoice_type')
                    if invoice_type is None:
                        invoice_type = '其他发票'
                    
                    # 确保 amount 不为 None
                    amount = data.get('amount', 0)
                    if amount is None:
                        amount = 0
                    
                    # 传入 is_duplicate 参数，重复发票不计入金额统计
                    type_folder = organizer.get_type_folder(
                        invoice_type, 
                        amount, 
                        dup_info.get('is_duplicate', False)
                    )
                    # 传入 is_duplicate 参数，重复发票不参与PDF合并
                    final_path = organizer.move_file(
                        file_path, 
                        type_folder, 
                        new_name, 
                        is_duplicate=dup_info.get('is_duplicate', False)
                    )
                    
                    # 更新记录 - 确保所有字段都有值
                    record.invoice_num = data.get('invoice_num_full', '') or ''
                    record.amount = data.get('amount', 0) or 0
                    record.invoice_type = invoice_type or '其他发票'
                    record.invoice_date = (data.get('date', '') or '')[:7]
                    record.seller_name = data.get('seller_name', '') or ''
                    record.buyer_name = data.get('buyer_name', '') or ''  # 添加购买方名称，确保不为None
                    record.is_success = True
                    record.is_duplicate = dup_info.get('is_duplicate', False)
                    record.output_path = final_path or ''
                    
                    result = {
                        'success': True,
                        'file_path': file_path,
                        'new_filename': new_name,
                        'final_path': final_path,
                        'is_duplicate': dup_info.get('is_duplicate', False),
                        'dup_index': dup_info.get('index', 0),
                        **data
                    }
                    results.append(result)
                    self.file_done.emit(result)
                    
                except Exception as e:
                    error_msg = str(e)
                    process_error = error_msg
                    record.error_msg = error_msg
                    record.is_success = False
                    
                    result = {
                        'success': False,
                        'file_path': file_path,
                        'error': error_msg,
                        'is_not_invoice': '不是发票' in error_msg
                    }
                    results.append(result)
                    self.file_done.emit(result)
                
                # 保存记录 - 独立的异常处理，确保不会卡住
                try:
                    self.stats_manager.add_record(record)
                except Exception as e:
                    error_detail = str(e)
                    print(f"保存记录失败 {file_path}: {error_detail}")
                    # 如果保存记录失败，尝试更新结果中的错误信息
                    if results:
                        results[-1]['error'] = results[-1].get('error', '') + f" [保存失败: {error_detail}]"
            
            # 最终处理
            try:
                organizer.finalize_folders()
            except Exception as e:
                print(f"finalize_folders 失败: {e}")
            
            # 计算汇总（剔除重复发票的金额，保留两位小数）
            try:
                success_results = [r for r in results if r.get('success')]
                # 只统计非重复发票的金额
                non_duplicate_results = [r for r in success_results if not r.get('is_duplicate', False)]
                total_amount = round(sum(r.get('amount', 0) or 0 for r in non_duplicate_results), 2)
                
                # 保存任务汇总
                self.stats_manager.save_task_summary(
                    total=total,
                    success=len(success_results),
                    failed=total - len(success_results),
                    duplicate=sum(1 for r in success_results if r.get('is_duplicate')),
                    total_amount=total_amount,
                    output_folder=str(organizer.current_task_folder)
                )
                
                # 生成Excel（返回路径、总金额、类型汇总）
                excel_result = excel.generate(success_results, organizer.current_task_folder)
                
                # 确保返回值是三元组
                if excel_result is None:
                    excel_path = None
                    total_amount_excel = 0
                    type_summary = {}
                elif isinstance(excel_result, tuple):
                    if len(excel_result) >= 3:
                        excel_path, total_amount_excel, type_summary = excel_result[0], excel_result[1], excel_result[2]
                    elif len(excel_result) == 2:
                        excel_path, total_amount_excel = excel_result
                        type_summary = {}
                    else:
                        excel_path = excel_result[0]
                        total_amount_excel = 0
                        type_summary = {}
                else:
                    excel_path = excel_result
                    total_amount_excel = 0
                    type_summary = {}
            except Exception as e:
                print(f"生成汇总或Excel失败: {e}")
                excel_path = None
                total_amount = 0
                type_summary = {}
            
            # 清理解压的临时文件（保留原压缩包）
            for file_path in self.temp_extracted_files:
                try:
                    if os.path.exists(file_path):
                        os.remove(file_path)
                except:
                    pass
            
            # 尝试删除临时目录
            try:
                temp_dir = Path(self.output_path) / f"_temp_extracted_{datetime.now().strftime('%H%M%S')}"
                if temp_dir.exists():
                    temp_dir.rmdir()
            except:
                pass
            
            self.finished_ok.emit({
                'total': total,
                'success': len([r for r in results if r.get('success')]),
                'failed': len([r for r in results if not r.get('success')]),
                'excel_path': excel_path,
                'output_folder': str(organizer.current_task_folder) if hasattr(organizer, 'current_task_folder') else self.output_path,
                'total_amount': total_amount if 'total_amount' in locals() else 0,
                'type_summary': type_summary if 'type_summary' in locals() and type_summary else {}
            })
            
        except Exception as e:
            import traceback
            error_detail = f"{str(e)}\n{traceback.format_exc()}"
            print(f"WorkerThread 致命错误: {error_detail}")
            self.error.emit(f"处理失败: {str(e)}")
    
    def stop(self):
        self.running = False


class StatisticsDialog(QDialog):
    """统计查询对话框"""
    
    def __init__(self, stats_manager: StatisticsManager, parent=None):
        super().__init__(parent)
        self.stats_manager = stats_manager
        self.setWindowTitle("统计查询与导出")
        self.setGeometry(100, 100, 800, 600)
        
        self.init_ui()
        self.load_history()
    
    def init_ui(self):
        layout = QVBoxLayout(self)
        
        # 标签页
        tabs = QTabWidget()
        
        # 1. 当前任务统计
        tab_current = QWidget()
        current_layout = QVBoxLayout(tab_current)
        
        self.lbl_current_stats = QLabel("暂无数据")
        self.lbl_current_stats.setStyleSheet("font-size: 14px; padding: 10px;")
        current_layout.addWidget(self.lbl_current_stats)
        
        # 实时统计表格
        self.tree_type = QTreeWidget()
        self.tree_type.setHeaderLabels(['发票类型', '数量', '金额'])
        current_layout.addWidget(QLabel("按类型统计:"))
        current_layout.addWidget(self.tree_type)
        
        self.tree_month = QTreeWidget()
        self.tree_month.setHeaderLabels(['年月', '数量', '金额'])
        current_layout.addWidget(QLabel("按年月统计:"))
        current_layout.addWidget(self.tree_month)
        
        tabs.addTab(tab_current, "当前任务")
        
        # 2. 历史记录
        tab_history = QWidget()
        history_layout = QVBoxLayout(tab_history)
        
        self.tree_history = QTreeWidget()
        self.tree_history.setHeaderLabels(['任务ID', '日期', '总数', '成功', '失败', '重复', '总金额', '操作'])
        self.tree_history.itemClicked.connect(self.on_history_selected)
        history_layout.addWidget(self.tree_history)
        
        btn_refresh = QPushButton("刷新历史")
        btn_refresh.clicked.connect(self.load_history)
        history_layout.addWidget(btn_refresh)
        
        tabs.addTab(tab_history, "历史记录")
        
        layout.addWidget(tabs)
        
        # 导出按钮
        btn_layout = QHBoxLayout()
        
        btn_export_current = QPushButton("导出当前统计")
        btn_export_current.clicked.connect(self.export_current)
        btn_layout.addWidget(btn_export_current)
        
        btn_export_history = QPushButton("导出选中历史")
        btn_export_history.clicked.connect(self.export_history)
        btn_layout.addWidget(btn_export_history)
        
        btn_layout.addStretch()
        
        btn_close = QPushButton("关闭")
        btn_close.clicked.connect(self.accept)
        btn_layout.addWidget(btn_close)
        
        layout.addLayout(btn_layout)
        
        self.selected_history_task = None
    
    def update_current_stats(self):
        """更新当前统计显示"""
        stats = self.stats_manager.get_current_statistics()
        
        # 确保 stats 有默认值
        if stats is None:
            stats = {
                'task_id': '',
                'total': 0,
                'success': 0,
                'failed': 0,
                'duplicate': 0,
                'type_statistics': {},
                'month_statistics': {}
            }
        
        text = f"""
        <b>任务ID:</b> {stats.get('task_id', '')}<br>
        <b>总计:</b> {stats.get('total', 0)} 张<br>
        <b style="color: green;">成功:</b> {stats.get('success', 0)} 张<br>
        <b style="color: red;">失败:</b> {stats.get('failed', 0)} 张<br>
        <b style="color: orange;">重复:</b> {stats.get('duplicate', 0)} 张
        """
        self.lbl_current_stats.setText(text)
        
        # 更新类型统计
        self.tree_type.clear()
        type_statistics = stats.get('type_statistics', {})
        if type_statistics:
            for type_name, data in type_statistics.items():
                item = QTreeWidgetItem([
                    type_name,
                    str(data.get('count', 0)),
                    f"{data.get('amount', 0):.2f}"
                ])
                self.tree_type.addTopLevelItem(item)
        
        # 更新年月统计
        self.tree_month.clear()
        month_statistics = stats.get('month_statistics', {})
        if month_statistics:
            for month, data in sorted(month_statistics.items()):
                item = QTreeWidgetItem([
                    month,
                    str(data.get('count', 0)),
                    f"{data.get('amount', 0):.2f}"
                ])
                self.tree_month.addTopLevelItem(item)
    
    def load_history(self):
        """加载历史记录"""
        self.tree_history.clear()
        tasks = self.stats_manager.get_history_tasks(20)
        
        if tasks:
            for task in tasks:
                item = QTreeWidgetItem([
                    task.get('task_id', ''),
                    task.get('process_date', ''),
                    str(task.get('total_count', 0)),
                    str(task.get('success_count', 0)),
                    str(task.get('failed_count', 0)),
                    str(task.get('duplicate_count', 0)),
                    f"{task.get('total_amount', 0):.2f}",
                    "点击选择"
                ])
                item.task_id = task.get('task_id')
                self.tree_history.addTopLevelItem(item)
    
    def on_history_selected(self, item):
        """选择历史记录"""
        if hasattr(item, 'task_id'):
            self.selected_history_task = item.task_id
            QMessageBox.information(self, "已选择", f"已选中任务: {item.task_id}")
    
    def export_current(self):
        """导出当前统计"""
        path = QFileDialog.getExistingDirectory(self, "选择导出位置")
        if path:
            try:
                file_path = self.stats_manager.export_to_excel(path)
                QMessageBox.information(self, "成功", f"已导出到:\n{file_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", str(e))
    
    def export_history(self):
        """导出历史记录"""
        if not self.selected_history_task:
            QMessageBox.warning(self, "提示", "请先点击选择一条历史记录")
            return
        
        path = QFileDialog.getExistingDirectory(self, "选择导出位置")
        if path:
            try:
                file_path = self.stats_manager.export_to_excel(path, self.selected_history_task)
                QMessageBox.information(self, "成功", f"已导出到:\n{file_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", str(e))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("发票智能整理工具 v2.0")
        self.setGeometry(100, 100, 1000, 700)
        
        self.config = ConfigManager()
        self.logger = Logger()
        self.stats_manager = StatisticsManager()
        
        self.worker = None
        self.last_output_folder = None  # 保存上次输出文件夹
        self.last_type_summary = None   # 保存类型汇总
        
        self.init_ui()
        self.load_config()
        self.init_menu()
    
    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        
        # 启用拖拽支持
        self.setAcceptDrops(True)
        
        # 主分割器
        splitter = QSplitter(Qt.Horizontal)
        
        # 左边：主要操作区
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        
        # 路径设置
        path_group = QGroupBox("路径设置（支持拖拽文件夹到窗口）")
        path_layout = QVBoxLayout()
        
        h1 = QHBoxLayout()
        h1.addWidget(QLabel("发票来源:"))
        self.input_edit = QLineEdit()
        self.input_edit.setReadOnly(True)
        self.input_edit.setPlaceholderText("点击浏览选择文件夹，或直接拖拽文件夹到窗口")
        h1.addWidget(self.input_edit)
        btn_browse = QPushButton("浏览...")
        btn_browse.setToolTip("点击打开自定义文件夹选择器，可以查看文件但只能选择文件夹")
        btn_browse.clicked.connect(self.select_input_folder)
        h1.addWidget(btn_browse)
        path_layout.addLayout(h1)
        
        h2 = QHBoxLayout()
        h2.addWidget(QLabel("输出位置:"))
        self.output_edit = QLineEdit()
        h2.addWidget(self.output_edit)
        btn_out = QPushButton("浏览...")
        btn_out.clicked.connect(self.select_output)
        h2.addWidget(btn_out)
        path_layout.addLayout(h2)
        
        path_group.setLayout(path_layout)
        left_layout.addWidget(path_group)
        
        # 实时统计面板
        stats_group = QGroupBox("实时统计")
        stats_layout = QVBoxLayout()
        
        self.lbl_stats = QLabel("等待开始...")
        self.lbl_stats.setStyleSheet("""
            QLabel {
                background-color: #f0f0f0;
                padding: 10px;
                border-radius: 5px;
                font-size: 12px;
            }
        """)
        stats_layout.addWidget(self.lbl_stats)
        
        stats_group.setLayout(stats_layout)
        left_layout.addWidget(stats_group)
        
        # 文件列表
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["文件名", "状态", "发票号", "金额", "备注"])
        self.table.setMaximumHeight(200)
        left_layout.addWidget(self.table)
        
        # 日志
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        left_layout.addWidget(self.log)
        
        # 进度和控制
        self.progress = QProgressBar()
        left_layout.addWidget(self.progress)
        
        self.status_label = QLabel("就绪")
        left_layout.addWidget(self.status_label)
        
        btn_layout = QHBoxLayout()
        
        self.btn_start = QPushButton("开始整理")
        self.btn_start.setStyleSheet("background-color: #4CAF50; color: white; padding: 10px; font-weight: bold;")
        self.btn_start.clicked.connect(self.start)
        btn_layout.addWidget(self.btn_start)
        
        self.btn_stop = QPushButton("停止")
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self.stop)
        btn_layout.addWidget(self.btn_stop)
        
        btn_layout.addStretch()
        
        btn_config = QPushButton("配置API")
        btn_config.clicked.connect(self.config_api)
        btn_layout.addWidget(btn_config)
        
        left_layout.addLayout(btn_layout)
        
        splitter.addWidget(left_widget)
        
        # 右边：统计面板
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        
        stats_title = QLabel("统计面板")
        stats_title.setStyleSheet("font-size: 16px; font-weight: bold; padding: 10px;")
        right_layout.addWidget(stats_title)
        
        # 当前任务统计
        self.tree_current_type = QTreeWidget()
        self.tree_current_type.setHeaderLabels(['发票类型', '数量', '金额'])
        right_layout.addWidget(QLabel("按类型统计:"))
        right_layout.addWidget(self.tree_current_type)
        
        # 导出按钮
        btn_export = QPushButton("导出详细统计报表")
        btn_export.clicked.connect(self.show_statistics)
        right_layout.addWidget(btn_export)
        
        right_layout.addStretch()
        
        splitter.addWidget(right_widget)
        splitter.setSizes([600, 400])
        
        # 设置主布局
        main_layout = QVBoxLayout(central)
        main_layout.addWidget(splitter)
    
    def init_menu(self):
        """初始化菜单"""
        menubar = self.menuBar()
        
        # 统计菜单
        stats_menu = menubar.addMenu("统计")
        
        action_stats = QAction("查看统计", self)
        action_stats.triggered.connect(self.show_statistics)
        stats_menu.addAction(action_stats)
        
        action_history = QAction("历史记录", self)
        action_history.triggered.connect(self.show_history)
        stats_menu.addAction(action_history)
        
        # 帮助菜单
        help_menu = menubar.addMenu("帮助")
        
        action_about = QAction("关于", self)
        action_about.triggered.connect(self.show_about)
        help_menu.addAction(action_about)
    
    def load_config(self):
        self.input_edit.setText(self.config.get_last_path('input'))
        self.output_edit.setText(self.config.get_last_path('output') or str(Path.home() / '发票整理'))
    
    def select_input_folder(self):
        """使用自定义对话框选择文件夹"""
        initial = self.input_edit.text() or self.config.get_last_path('input')
        
        # 使用自定义对话框
        dialog = FolderDialog(self, initial)
        result = dialog.exec_()
        
        if result == QDialog.Accepted:
            folder = dialog.get_selected_folder()
            if folder:
                self._set_input_folder(folder)
    
    def select_output(self):
        path = QFileDialog.getExistingDirectory(self, "选择输出位置", self.output_edit.text())
        if path:
            self.output_edit.setText(path)
            self.config.set_last_path('output', path)
    
    def config_api(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("配置百度OCR API")
        layout = QFormLayout(dialog)
        
        from PyQt5.QtWidgets import QLineEdit as QLE
        le_appid = QLE()
        le_apikey = QLE()
        le_secret = QLE()
        
        creds = self.config.get_api_credentials()
        le_appid.setText(creds.get('app_id', ''))
        le_apikey.setText(creds.get('api_key', ''))
        le_secret.setText(creds.get('secret_key', ''))
        
        layout.addRow("App ID:", le_appid)
        layout.addRow("API Key:", le_apikey)
        layout.addRow("Secret Key:", le_secret)
        
        btn_ok = QPushButton("保存")
        btn_ok.clicked.connect(dialog.accept)
        layout.addRow(btn_ok)
        
        if dialog.exec_():
            self.config.set_api_credentials(le_appid.text(), le_apikey.text(), le_secret.text())
            self.config.save_config()
            QMessageBox.information(self, "提示", "配置已保存")
    
    def start(self):
        input_path = self.input_edit.text()
        output_path = self.output_edit.text()
        
        if not input_path:
            QMessageBox.warning(self, "警告", "请先选择发票文件夹！")
            return
        
        if Path(input_path).is_file():
            QMessageBox.warning(
                self, 
                "错误", 
                "只能选择文件夹，不能选择具体文件！\n\n"
                "系统会自动整理该文件夹下的所有发票文件（PDF、JPG等）\n"
                "请重新点击\"浏览...\"按钮选择文件夹"
            )
            return
        
        if not Path(input_path).exists():
            QMessageBox.warning(self, "警告", "选择的文件夹不存在")
            return
        
        if not output_path:
            QMessageBox.warning(self, "警告", "请选择输出位置")
            return
        
        creds = self.config.get_api_credentials()
        if not all([creds.get('app_id'), creds.get('api_key'), creds.get('secret_key')]):
            QMessageBox.warning(self, "警告", "请先配置百度OCR API")
            return
        
        # 清空
        self.table.setRowCount(0)
        self.log.clear()
        self.tree_current_type.clear()
        
        # ===== 预处理压缩包 =====
        self.status_label.setText("正在扫描压缩包...")
        QApplication.processEvents()  # 刷新UI
        
        from archive_handler import ArchiveHandler, scan_and_collect_files
        from gui.password_dialog import ArchivePasswordDialog
        
        archive_handler = ArchiveHandler(self)
        ArchivePasswordDialog.reset_session()  # 重置密码缓存
        
        # 扫描文件和压缩包
        normal_files, archive_infos = scan_and_collect_files(input_path, archive_handler, self)
        
        # 处理需要密码的压缩包
        temp_extracted_files = []
        processed_archives = []
        
        for idx, archive_info in enumerate(archive_infos):
            if archive_info.get('skip'):
                self.log.append(f'<span style="color: orange;">跳过压缩包: {archive_info.get("name", "")} (密码跳过)</span>')
                continue
            
            if archive_info.get('is_password_protected') and not archive_info.get('password'):
                # 需要密码但未输入，跳过
                continue
            
            # 解压到临时目录 - 使用唯一命名避免冲突
            temp_dir = Path(output_path) / f"_temp_extracted_{datetime.now().strftime('%H%M%S')}_{idx}"
            temp_dir.mkdir(parents=True, exist_ok=True)
            
            self.status_label.setText(f"正在解压: {archive_info.get('name', '')}...")
            QApplication.processEvents()
            
            extracted, status = archive_handler.extract_archive(
                archive_info.get('path'),
                str(temp_dir),
                archive_info.get('password'),
                force=True  # 强制重新解压，因为每次整理都应该解压到新位置
            )
            
            if status == 'success':
                # 使用 extract_archive 返回的文件列表
                if extracted:
                    temp_extracted_files.extend(extracted)
                    processed_archives.append({
                        'name': archive_info.get('name', ''),
                        'extracted_count': len(extracted)
                    })
                    self.log.append(f'<span style="color: green;">解压完成: {archive_info.get("name", "")} ({len(extracted)}个文件)</span>')
                else:
                    self.log.append(f'<span style="color: orange;">压缩包内无发票文件: {archive_info.get("name", "")}</span>')
                    # 清理空临时目录
                    try:
                        temp_dir.rmdir()
                    except:
                        pass
            elif status == 'already_extracted':
                # 已经解压过，跳过
                self.log.append(f'<span style="color: orange;">已处理过: {archive_info.get("name", "")}</span>')
                # 清理临时目录
                try:
                    temp_dir.rmdir()
                except:
                    pass
            elif status == 'no_invoice':
                self.log.append(f'<span style="color: orange;">压缩包内无发票: {archive_info.get("name", "")}</span>')
                # 清理临时目录
                try:
                    temp_dir.rmdir()
                except:
                    pass
            elif status.startswith('error'):
                self.log.append(f'<span style="color: red;">解压失败 {archive_info.get("name", "")}: {status}</span>')
                # 清理临时目录
                try:
                    temp_dir.rmdir()
                except:
                    pass
        
        # 合并文件列表
        all_files = list(normal_files) + temp_extracted_files
        
        if len(all_files) == 0:
            # 清理临时文件
            archive_handler.cleanup_extracted_files(temp_extracted_files)
            QMessageBox.information(self, "提示", "未找到有效的发票文件")
            return
        
        self.log.append(f'<span style="color: blue;">共找到 {len(normal_files)} 个普通文件，{len(temp_extracted_files)} 个压缩包内文件</span>')
        
        # 启动工作线程
        self.worker = WorkerThread(input_path, output_path, self.stats_manager, all_files, temp_extracted_files)
        self.worker.progress_updated.connect(self.update_progress)
        self.worker.file_done.connect(self.on_file_done)
        self.worker.finished_ok.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()
        
        self.btn_start.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.status_label.setText(f"正在处理 {len(all_files)} 个文件...")
    
    def stop(self):
        if self.worker:
            self.worker.stop()
            self.worker.wait()
            self.reset_ui()
    
    def update_progress(self, current, total, filename):
        self.progress.setMaximum(total)
        self.progress.setValue(current)
        self.status_label.setText(f"处理中: {filename} ({current}/{total})")
        
        # 更新实时统计
        stats = self.stats_manager.get_current_statistics()
        if stats:
            text = f"总计: {stats.get('total', 0)} | 成功: {stats.get('success', 0)} | 失败: {stats.get('failed', 0)} | 重复: {stats.get('duplicate', 0)}"
            self.lbl_stats.setText(text)
    
    def on_file_done(self, result):
        row = self.table.rowCount()
        self.table.insertRow(row)
        
        filename = Path(result.get('file_path', '')).name
        self.table.setItem(row, 0, QTableWidgetItem(filename))
        
        if result.get('success'):
            status = "成功"
            if result.get('is_duplicate'):
                status += f" [重复{result.get('dup_index', 0)}]"
            self.table.setItem(row, 1, QTableWidgetItem(status))
            self.table.setItem(row, 2, QTableWidgetItem(result.get('invoice_num_full', '')))
            self.table.setItem(row, 3, QTableWidgetItem(str(result.get('amount_int', 0))))
            
            # 更新右侧统计树
            self.update_stats_tree()
            
            color = "green"
            log_msg = "成功"
            if result.get('is_duplicate'):
                log_msg += f" [发现重复发票 #{result.get('dup_index', 0)}]"
        else:
            # 区分"不是发票"和其他错误
            is_not_invoice = result.get('is_not_invoice', False)
            error_msg = result.get('error', '')
            
            if is_not_invoice:
                status_item = QTableWidgetItem("非发票")
                status_item.setForeground(QColor(255, 140, 0))  # 橙色
                self.table.setItem(row, 1, status_item)
                # 备注列显示详细提示
                remark_item = QTableWidgetItem("识别出来不是发票")
                remark_item.setForeground(QColor(255, 140, 0))
                self.table.setItem(row, 4, remark_item)
                color = "orange"
            else:
                self.table.setItem(row, 1, QTableWidgetItem("失败"))
                self.table.setItem(row, 4, QTableWidgetItem(error_msg[:30]))
                color = "red"
            
            log_msg = error_msg
        
        self.log.append(f'<span style="color: {color};">{filename}: {log_msg}</span>')
    
    def update_stats_tree(self):
        """更新统计树"""
        stats = self.stats_manager.get_current_statistics()
        
        if not stats:
            return
        
        # 类型统计
        self.tree_current_type.clear()
        type_statistics = stats.get('type_statistics', {})
        if type_statistics:
            for type_name, data in type_statistics.items():
                item = QTreeWidgetItem([
                    type_name,
                    str(data.get('count', 0)),
                    f"¥{data.get('amount', 0):.2f}"
                ])
                self.tree_current_type.addTopLevelItem(item)
    
    def on_finished(self, result):
        self.reset_ui()
        
        # 保存输出信息（用于打印）
        self.last_output_folder = result.get('output_folder')
        self.last_type_summary = result.get('type_summary', {})
        
        # 显示完成统计
        stats = self.stats_manager.get_current_statistics()
        duplicate_count = stats.get('duplicate', 0) if stats else 0
        
        # 创建自定义完成对话框
        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QTextEdit, QPushButton
        
        finish_dialog = QDialog(self)
        finish_dialog.setWindowTitle("处理完成")
        finish_dialog.setMinimumSize(500, 400)
        finish_dialog.setGeometry(200, 200, 550, 450)
        # 设置为应用程序模态对话框
        finish_dialog.setWindowModality(Qt.ApplicationModal)
        
        layout = QVBoxLayout(finish_dialog)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # 统计信息显示
        info_text = f"""
<h2 style='color: #4CAF50;'>发票整理完成！</h2>
<hr>
<table style='font-size: 14px;'>
<tr><td><b>总计:</b></td><td>{result.get('total', 0)} 张</td></tr>
<tr><td style='color: green;'><b>成功:</b></td><td>{result.get('success', 0)} 张</td></tr>
<tr><td style='color: red;'><b>失败:</b></td><td>{result.get('failed', 0)} 张</td></tr>
<tr><td style='color: orange;'><b>重复:</b></td><td>{duplicate_count} 张</td></tr>
</table>
<hr>
<p style='font-size: 12px; color: #666;'>
<b>Excel已保存至:</b><br>{result.get('excel_path', '')}
</p>
<p style='font-size: 12px; color: #666;'>
<b>输出文件夹:</b><br>{result.get('output_folder', '')}
</p>
"""
        
        info_label = QTextEdit()
        info_label.setHtml(info_text)
        info_label.setReadOnly(True)
        info_label.setMaximumHeight(200)
        info_label.setStyleSheet("""
            QTextEdit {
                border: none;
                background-color: #f9f9f9;
                border-radius: 8px;
                padding: 10px;
            }
        """)
        layout.addWidget(info_label)
        
        # 按钮区
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        # 打开文件夹按钮
        open_folder_btn = QPushButton("打开整理后文件夹")
        open_folder_btn.setStyleSheet("""
            QPushButton {
                padding: 12px 24px;
                font-size: 14px;
                background-color: #607D8B;
                color: white;
                border: none;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #455A64;
            }
        """)
        output_folder = result.get('output_folder', '')
        open_folder_btn.clicked.connect(lambda: self.open_output_folder(output_folder))
        button_layout.addWidget(open_folder_btn)
        
        # 统计按钮
        stats_btn = QPushButton("查看统计")
        stats_btn.setStyleSheet("""
            QPushButton {
                padding: 12px 24px;
                font-size: 14px;
                background-color: #2196F3;
                color: white;
                border: none;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
        """)
        stats_btn.clicked.connect(lambda: [finish_dialog.accept(), self.show_statistics()])
        button_layout.addWidget(stats_btn)
        
        # 打印按钮（重点）
        print_btn = QPushButton("打印合并文档")
        print_btn.setStyleSheet("""
            QPushButton {
                padding: 12px 24px;
                font-size: 14px;
                background-color: #FF9800;
                color: white;
                border: none;
                border-radius: 6px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #F57C00;
            }
        """)
        print_btn.clicked.connect(lambda: [finish_dialog.accept(), self.show_print_dialog()])
        button_layout.addWidget(print_btn)
        
        # 确定按钮
        ok_btn = QPushButton("确定")
        ok_btn.setStyleSheet("""
            QPushButton {
                padding: 12px 24px;
                font-size: 14px;
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        ok_btn.clicked.connect(finish_dialog.accept)
        ok_btn.setDefault(True)
        button_layout.addWidget(ok_btn)
        
        layout.addLayout(button_layout)
        
        finish_dialog.exec_()
    
    def open_output_folder(self, folder_path: str):
        """打开输出文件夹"""
        if folder_path and os.path.exists(folder_path):
            # 使用 QDesktopServices 打开文件夹
            url = QUrl.fromLocalFile(folder_path)
            QDesktopServices.openUrl(url)
        else:
            QMessageBox.warning(self, "提示", "文件夹不存在或路径无效")
    
    def show_print_dialog(self):
        """显示打印设置对话框"""
        from gui.print_dialog import PrintSettingsDialog
        
        # 使用保存的类型汇总（非重复的）
        type_stats = self.last_type_summary or {}
        
        # 获取输出文件夹
        output_folder = self.last_output_folder
        if not output_folder:
            QMessageBox.warning(self, "提示", "未找到输出文件夹")
            return
        
        # 显示打印对话框
        PrintSettingsDialog.show_print_dialog(type_stats, str(output_folder), self)
    
    def on_error(self, error_msg):
        self.reset_ui()
        QMessageBox.critical(self, "错误", error_msg)
    
    def reset_ui(self):
        self.btn_start.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.status_label.setText("就绪")
    
    def show_statistics(self):
        """显示统计对话框"""
        dialog = StatisticsDialog(self.stats_manager, self)
        dialog.update_current_stats()
        dialog.exec_()
    
    def show_history(self):
        """显示历史记录"""
        self.show_statistics()
    
    def show_about(self):
        QMessageBox.about(self, "关于", """
        <h2>发票智能整理工具 v2.0</h2>
        <p>功能特性：</p>
        <ul>
            <li>自定义文件夹选择器（可查看文件但只能选择文件夹）</li>
            <li>自动识别发票信息（百度OCR）</li>
            <li>智能查重和分类整理</li>
            <li>实时统计面板（按类型、按年月）</li>
            <li>历史记录永久保存</li>
            <li>一键导出统计报表</li>
        </ul>
        <p>数据保存在: ~/.invoice_processor/statistics/</p>
        """)
    
    def closeEvent(self, event):
        self.config.save_config()
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.worker.wait(2000)
        event.accept()
    
    # ========== 拖拽支持 ==========
    def dragEnterEvent(self, event):
        """拖拽进入事件"""
        if event.mimeData().hasUrls():
            # 检查是否为文件夹
            urls = event.mimeData().urls()
            if urls:
                path = urls[0].toLocalFile()
                if os.path.isdir(path):
                    event.acceptProposedAction()
                    self.status_label.setText("松开鼠标导入此文件夹")
                    return
        event.ignore()
    
    def dragLeaveEvent(self, event):
        """拖拽离开事件"""
        self.status_label.setText("就绪")
    
    def dropEvent(self, event):
        """拖拽放下事件"""
        urls = event.mimeData().urls()
        if urls:
            folder_path = urls[0].toLocalFile()
            if os.path.isdir(folder_path):
                self._set_input_folder(folder_path)
        event.acceptProposedAction()
    
    def _set_input_folder(self, folder_path: str):
        """设置输入文件夹并统计文件（递归统计所有子文件夹）"""
        self.input_edit.setText(folder_path)
        self.config.set_last_path('input', folder_path)
        
        # 递归统计文件数量（包括子文件夹）
        path = Path(folder_path)
        file_count = 0
        for ext in ['.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif']:
            # 使用 rglob 递归搜索所有子文件夹
            file_count += len(list(path.rglob(f'*{ext}')))
            file_count += len(list(path.rglob(f'*{ext.upper()}')))
        
        self.status_label.setText(f"已选择: {path.name} (包含{file_count}个发票文件)")


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
