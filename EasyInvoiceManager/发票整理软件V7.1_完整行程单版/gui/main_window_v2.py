import hashlib
"""
主窗口 V2.0（支持铁路电子客票和双引擎OCR）
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
from ocr_engine_v2 import OCREngineV2
from data_parser_v2 import DataParserV2
from duplicate_checker import DuplicateChecker
from duplicate_checker import DuplicateChecker
from file_organizer_v2 import FileOrganizerV2
from excel_generator_v2 import ExcelGeneratorV2
from statistics_manager import StatisticsManager, InvoiceRecord
from utils.helpers import calculate_md5

from gui.folder_dialog import FolderDialog


class WorkerThreadV2(QThread):
    """工作线程 V2.0"""
    progress_updated = pyqtSignal(int, int, str)
    file_done = pyqtSignal(dict)
    finished_ok = pyqtSignal(dict)
    error = pyqtSignal(str)
    
    def __init__(self, input_path, output_path, stats_manager, 
                 files_to_process=None, temp_extracted_files=None):
        super().__init__()
        self.input_path = input_path
        self.output_path = output_path
        self.stats_manager = stats_manager
        self.files_to_process = files_to_process or []
        self.temp_extracted_files = temp_extracted_files or []
        self.running = True
    
    def run(self):
        try:
            # 使用新模块
            ocr = OCREngineV2()
            parser = DataParserV2()
            checker = DuplicateChecker()
            organizer = FileOrganizerV2(self.output_path)
            excel = ExcelGeneratorV2()
            
            organizer.create_folder_structure()
            
            files = self.files_to_process
            total = len(files)
            
            if total == 0:
                self.error.emit("未找到有效的发票文件")
                return
            
            # 开始统计任务
            self.stats_manager.start_task()
            
            # 分类存储结果
            vat_results = []       # 增值税发票
            railway_results = []   # 铁路电子客票
            itinerary_results = [] # 行程单
            failed_list = []       # 识别失败
            
            for i, file_path in enumerate(files, 1):
                if not self.running:
                    break
                
                self.progress_updated.emit(i, total, Path(file_path).name)
                
                # 处理单个文件
                result = self._process_single_file(
                    file_path, ocr, parser, checker, organizer
                )
                
                # 分类结果（result是列表）
                for item in result:
                    if item.get('is_failed'):
                        failed_list.append({
                            'filename': Path(file_path).name,
                            'reason': item.get('error', '未知错误'),
                            'file_path': file_path
                        })
                    elif item.get('is_railway'):
                        railway_results.append(item)
                    elif item.get('is_itinerary'):
                        itinerary_results.append(item)
                    else:
                        vat_results.append(item)
                
                # 逐个发送信号（支持列表）
                    for item in (result if isinstance(result, list) else [result]):
                        self.file_done.emit(item)
            
            # 最终处理
            try:
                organizer.finalize_folders()
            except Exception as e:
                print(f"finalize_folders 失败: {e}")
            
            # 生成Excel
            try:
                excel_result = excel.generate(
                    vat_results, railway_results, itinerary_results, failed_list,
                    organizer.current_task_folder
                )
                excel_path = excel_result[0] if excel_result else None
                total_amount = excel_result[1] if excel_result else 0
                type_summary = excel_result[2] if excel_result else {}
            except Exception as e:
                print(f"生成Excel失败: {e}")
                excel_path = None
                total_amount = 0
                type_summary = {}
            
            # 保存任务汇总
            success_count = len([r for r in vat_results + railway_results + itinerary_results if r.get('success')])
            duplicate_count = sum(1 for r in vat_results + railway_results + itinerary_results if r.get('is_duplicate'))
            
            self.stats_manager.save_task_summary(
                total=total,
                success=success_count,
                failed=len(failed_list),
                duplicate=duplicate_count,
                total_amount=total_amount,
                output_folder=str(organizer.current_task_folder)
            )
            
            # 清理解压的临时文件
            for file_path in self.temp_extracted_files:
                try:
                    if os.path.exists(file_path):
                        os.remove(file_path)
                except:
                    pass
            
            self.finished_ok.emit({
                'total': total,
                'success': success_count,
                'failed': len(failed_list),
                'excel_path': excel_path,
                'output_folder': str(organizer.current_task_folder),
                'total_amount': total_amount,
                'type_summary': type_summary,
                'itinerary_count': len(itinerary_results)
            })
            
        except Exception as e:
            import traceback
            error_detail = f"{str(e)}\n{traceback.format_exc()}"
            print(f"WorkerThread 致命错误: {error_detail}")
            self.error.emit(f"处理失败: {str(e)}")
    
    def _process_single_file(self, file_path: str, ocr, parser, checker, organizer) -> list:
        """处理单个文件 - 返回列表支持多段行程展开"""
        filename = Path(file_path).name
        
        # 创建记录
        record = InvoiceRecord(
            task_id=self.stats_manager.current_task_id,
            file_name=filename,
            file_path=file_path,
            invoice_num='',
            amount=0,
            invoice_type='',
            invoice_date='',
            seller_name='',
            buyer_name='',
            is_success=False,
            is_duplicate=False,
            error_msg='',
            process_time=datetime.now().isoformat(),
            output_path=''
        )
        
        try:
            # OCR识别
            ocr_result = ocr.recognize(file_path)
            
            # 解析数据（现在返回列表）
            data_list = parser.parse(ocr_result)
            
            # 确保是列表
            if not isinstance(data_list, list):
                data_list = [data_list]
            
            results = []
            
            for data in data_list:
                # 检查是否为铁路电子客票或行程单
                is_railway = data.get('is_railway', False)
                is_itinerary = data.get('is_itinerary', False)
                
                # 验证有效性 - 行程单跳过详细检查
                if not data.get('is_valid') and not is_itinerary:
                    if is_railway:
                        # 关键字段缺失，记录到失败列表
                        missing = data.get('missing_fields', [])
                        error_msg = f"关键字段缺失: {', '.join(missing)}"
                        
                        # 复制到失败文件夹
                        organizer.move_failed_file(file_path, filename)
                        
                        record.error_msg = error_msg
                        record.is_success = False
                        self.stats_manager.add_record(record)
                        
                        results.append({
                            'success': False,
                            'is_failed': True,
                            'is_railway': False,
                            'is_itinerary': False,
                            'file_path': file_path,
                            'error': error_msg,
                            'filename': filename
                        })
                        continue
                    else:
                        # 不是发票类型，跳过
                        continue
                
                # 对于行程单，强制设置为有效
                if is_itinerary:
                    data['is_valid'] = True
                
                # 检查重复（行程单按 日期+起点+终点+金额 去重）
                if is_itinerary:
                    # 构建行程单唯一键，直接计算字符串MD5
                    dup_key = f"{data.get('date', '')}_{data.get('departure', '')}_{data.get('arrival', '')}_{data.get('amount', 0)}"
                    dup_key_md5 = hashlib.md5(dup_key.encode('utf-8')).hexdigest()
                    dup_info = checker.check_duplicate({'invoice_num': dup_key}, dup_key_md5)
                else:
                    file_md5 = calculate_md5(file_path)
                    dup_info = checker.check_duplicate(data, file_md5)
                
                # 确定发票类型
                invoice_type = data.get('invoice_type', '其他发票')
                
                # 生成文件名
                ext = Path(file_path).suffix
                new_name = organizer.generate_filename(data, dup_info, ext, invoice_type)
                
                # 获取目标文件夹
                amount = data.get('amount', 0) or 0
                type_folder = organizer.get_type_folder(
                    invoice_type, amount, dup_info.get('is_duplicate', False)
                )
                
                # 移动文件（行程单多段只复制一次）
                if is_itinerary and data.get('segment_index', 1) > 1:
                    # 多段行程的后续段不复制文件，使用第一段的路径
                    final_path = results[0].get('final_path', '') if results else ''
                else:
                    final_path = organizer.move_file(
                        file_path, type_folder, new_name,
                        is_duplicate=dup_info.get('is_duplicate', False)
                    )
                
                # 更新记录
                record.invoice_num = data.get('invoice_num_full', '') or ''
                record.amount = amount
                record.invoice_type = invoice_type
                record.invoice_date = data.get('date', '') or ''
                record.seller_name = data.get('seller_name', '') or ''
                record.buyer_name = data.get('buyer_name', '') or data.get('passenger_name', '')
                record.is_success = True
                record.is_duplicate = dup_info.get('is_duplicate', False)
                record.output_path = final_path or ''
                
                self.stats_manager.add_record(record)
                
                # 构建结果
                result = {
                    'success': True,
                    'is_failed': False,
                    'is_railway': is_railway,
                    'is_itinerary': is_itinerary,
                    'file_path': file_path,
                    'new_filename': new_name,
                    'final_path': final_path,
                    'is_duplicate': dup_info.get('is_duplicate', False),
                    'dup_index': dup_info.get('index', 0),
                    **data
                }
                
                results.append(result)
            
            return results
            
        except Exception as e:
            error_msg = str(e)
            record.error_msg = error_msg
            record.is_success = False
            
            try:
                self.stats_manager.add_record(record)
            except:
                pass
            
            # 复制到失败文件夹
            try:
                organizer.move_failed_file(file_path, filename)
            except:
                pass
            
            return [{
                'success': False,
                'is_failed': True,
                'is_railway': False,
                'is_itinerary': False,
                'file_path': file_path,
                'error': error_msg,
                'filename': filename,
                'is_not_invoice': '不是发票' in error_msg
            }]
    
    def stop(self):
        self.running = False


class MainWindowV2(QMainWindow):
    """主窗口 V2.0"""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("发票智能整理工具 V7.0 - 纯API版（支持铁路电子客票）")
        self.setGeometry(100, 100, 1000, 700)
        
        self.config = ConfigManager()
        self.logger = Logger()
        self.stats_manager = StatisticsManager()
        
        self.worker = None
        self.last_output_folder = None
        self.last_type_summary = None
        
        self.init_ui()
        self.load_config()
        self.init_menu()
        self.check_ocr_engine()
    
    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        
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
        
        # OCR引擎状态显示
        ocr_group = QGroupBox("OCR引擎状态")
        ocr_layout = QVBoxLayout()
        self.lbl_ocr_status = QLabel("检测中...")
        ocr_layout.addWidget(self.lbl_ocr_status)
        ocr_group.setLayout(ocr_layout)
        left_layout.addWidget(ocr_group)
        
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
        self.table.setHorizontalHeaderLabels(["文件名", "状态", "发票号/类型", "金额", "备注"])
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
    
    def check_ocr_engine(self):
        """检查OCR引擎状态（纯API版本）"""
        try:
            from ocr_engine_v2 import OCREngineV2
            ocr = OCREngineV2()
            
            if ocr.baidu_client:
                self.lbl_ocr_status.setText("[OK] 百度API已配置（纯API模式）")
                self.lbl_ocr_status.setStyleSheet("color: green;")
            else:
                self.lbl_ocr_status.setText("[!] 百度API未配置，请先配置API密钥")
                self.lbl_ocr_status.setStyleSheet("color: red;")
                
        except Exception as e:
            error_msg = str(e)
            if "未配置" in error_msg or "未初始化" in error_msg:
                self.lbl_ocr_status.setText("[!] 百度API未配置，请先配置API密钥")
                self.lbl_ocr_status.setStyleSheet("color: red;")
            else:
                self.lbl_ocr_status.setText(f"[X] OCR引擎错误: {error_msg}")
                self.lbl_ocr_status.setStyleSheet("color: red;")
    
    def init_menu(self):
        menubar = self.menuBar()
        
        stats_menu = menubar.addMenu("统计")
        
        action_stats = QAction("查看统计", self)
        action_stats.triggered.connect(self.show_statistics)
        stats_menu.addAction(action_stats)
        
        action_history = QAction("历史记录", self)
        action_history.triggered.connect(self.show_history)
        stats_menu.addAction(action_history)
        
        help_menu = menubar.addMenu("帮助")
        
        action_about = QAction("关于", self)
        action_about.triggered.connect(self.show_about)
        help_menu.addAction(action_about)
    
    def load_config(self):
        self.input_edit.setText(self.config.get_last_path('input'))
        self.output_edit.setText(self.config.get_last_path('output') or str(Path.home() / '发票整理'))
    
    def select_input_folder(self):
        initial = self.input_edit.text() or self.config.get_last_path('input')
        dialog = FolderDialog(self, initial)
        result = dialog.exec_()
        
        if result == QDialog.Accepted:
            folder = dialog.get_selected_folder()
            if folder:
                self._set_input_folder(folder)
    
    def _set_input_folder(self, folder: str):
        self.input_edit.setText(folder)
        self.config.set_last_path('input', folder)
    
    def select_output(self):
        path = QFileDialog.getExistingDirectory(self, "选择输出位置", self.output_edit.text())
        if path:
            self.output_edit.setText(path)
            self.config.set_last_path('output', path)
    
    def config_api(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("配置百度OCR API（必须）")
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
        
        info_label = QLabel("【重要】本版本为纯API模式，必须配置百度API才能使用！\n\n申请方法：\n1. 访问 https://ai.baidu.com/tech/ocr\n2. 登录百度账号（无账号需注册）\n3. 点击'立即使用' → 创建应用\n4. 获取 App ID、API Key、Secret Key\n5. 将三个密钥填入上方输入框\n\n免费额度：每月1000次调用，足够一般使用")
        info_label.setStyleSheet("color: blue; font-size: 16px; font-weight: bold; line-height: 1.5;")
        layout.addRow(info_label)
        
        btn_ok = QPushButton("保存")
        btn_ok.clicked.connect(dialog.accept)
        layout.addRow(btn_ok)
        
        if dialog.exec_():
            self.config.set_api_credentials(le_appid.text(), le_apikey.text(), le_secret.text())
            self.config.save_config()
            self.check_ocr_engine()  # 刷新状态显示
            QMessageBox.information(self, "提示", "配置已保存")
    
    def start(self):
        input_path = self.input_edit.text()
        output_path = self.output_edit.text()
        
        if not input_path:
            QMessageBox.warning(self, "警告", "请先选择发票文件夹！")
            return
        
        if Path(input_path).is_file():
            QMessageBox.warning(
                self, "错误",
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
        
        # 检查API配置（纯API模式必须配置）
        creds = self.config.get_api_credentials()
        if not all([creds.get('app_id'), creds.get('api_key'), creds.get('secret_key')]):
            QMessageBox.warning(
                self, "警告",
                "请先配置百度OCR API密钥！\n\n"
                "本版本为纯API模式，必须配置API才能使用。\n"
                "申请地址：https://ai.baidu.com/tech/ocr\n"
                "每天500次免费额度。"
            )
            return
        
        # 清空
        self.table.setRowCount(0)
        self.log.clear()
        self.tree_current_type.clear()
        
        # 预处理压缩包
        self.status_label.setText("正在扫描压缩包...")
        QApplication.processEvents()
        
        from archive_handler import ArchiveHandler, scan_and_collect_files
        from gui.password_dialog import ArchivePasswordDialog
        
        archive_handler = ArchiveHandler(self)
        ArchivePasswordDialog.reset_session()
        
        normal_files, archive_infos = scan_and_collect_files(input_path, archive_handler, self)
        
        temp_extracted_files = []
        
        for idx, archive_info in enumerate(archive_infos):
            if archive_info.get('skip'):
                self.log.append(f'<span style="color: orange;">跳过压缩包: {archive_info.get("name", "")}</span>')
                continue
            
            if archive_info.get('is_password_protected') and not archive_info.get('password'):
                continue
            
            temp_dir = Path(output_path) / f"_temp_extracted_{datetime.now().strftime('%H%M%S')}_{idx}"
            temp_dir.mkdir(parents=True, exist_ok=True)
            
            self.status_label.setText(f"正在解压: {archive_info.get('name', '')}...")
            QApplication.processEvents()
            
            extracted, status = archive_handler.extract_archive(
                archive_info.get('path'),
                str(temp_dir),
                archive_info.get('password'),
                force=True
            )
            
            if status == 'success' and extracted:
                temp_extracted_files.extend(extracted)
                self.log.append(f'<span style="color: green;">解压完成: {archive_info.get("name", "")} ({len(extracted)}个文件)</span>')
        
        all_files = list(normal_files) + temp_extracted_files
        
        if len(all_files) == 0:
            archive_handler.cleanup_extracted_files(temp_extracted_files)
            QMessageBox.information(self, "提示", "未找到有效的发票文件")
            return
        
        self.log.append(f'<span style="color: blue;">共找到 {len(normal_files)} 个普通文件，{len(temp_extracted_files)} 个压缩包内文件</span>')
        
        # 启动工作线程
        self.worker = WorkerThreadV2(input_path, output_path, self.stats_manager, all_files, temp_extracted_files)
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
        
        stats = self.stats_manager.get_current_statistics()
        if stats:
            text = f"总计: {stats.get('total', 0)} | 成功: {stats.get('success', 0)} | 失败: {stats.get('failed', 0)} | 重复: {stats.get('duplicate', 0)}"
            self.lbl_stats.setText(text)
    
    def on_file_done(self, result):
        row = self.table.rowCount()
        self.table.insertRow(row)
        
        filename = Path(result.get('file_path', '')).name
        self.table.setItem(row, 0, QTableWidgetItem(filename))
        
        if result.get('is_failed'):
            # 识别失败
            is_not_invoice = result.get('is_not_invoice', False)
            if is_not_invoice:
                status_item = QTableWidgetItem("非发票")
                status_item.setForeground(QColor(255, 140, 0))
            else:
                status_item = QTableWidgetItem("识别失败")
                status_item.setForeground(QColor(255, 0, 0))
            self.table.setItem(row, 1, status_item)
            self.table.setItem(row, 4, QTableWidgetItem(result.get('error', '')[:30]))
            color = "orange" if is_not_invoice else "red"
            
        elif result.get('success'):
            # 识别成功
            if result.get('is_duplicate'):
                status = f"成功 [重复{result.get('dup_index', 0)}]"
                color = "orange"
            else:
                status = "成功"
                color = "green"
            
            self.table.setItem(row, 1, QTableWidgetItem(status))
            
            if result.get('is_railway'):
                # 铁路客票
                self.table.setItem(row, 2, QTableWidgetItem(f"铁路票 {result.get('departure', '')}-{result.get('arrival', '')}"))
            else:
                # 增值税发票
                self.table.setItem(row, 2, QTableWidgetItem(result.get('invoice_num_full', '')))
            
            self.table.setItem(row, 3, QTableWidgetItem(str(result.get('amount', 0))))
            self.update_stats_tree()
            
        self.log.append(f'<span style="color: {color};">{filename}: {result.get("error", "成功")}</span>')
    
    def update_stats_tree(self):
        stats = self.stats_manager.get_current_statistics()
        if not stats:
            return
        
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
        
        # 添加行程单统计（如果存在）
        # 注意：实际统计应该在statistics_manager中处理
    
    def on_finished(self, result):
        self.reset_ui()
        
        # 保存输出信息（用于打印和打开文件夹）
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
        
        # 查看Excel按钮
        excel_btn = QPushButton("打开Excel清单")
        excel_btn.setStyleSheet("""
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
        excel_path = result.get('excel_path', '')
        excel_btn.clicked.connect(lambda: self.open_excel_file(excel_path))
        button_layout.addWidget(excel_btn)
        
        # 打印设置按钮
        print_btn = QPushButton("打印设置")
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
            url = QUrl.fromLocalFile(folder_path)
            QDesktopServices.openUrl(url)
        else:
            QMessageBox.warning(self, "提示", "文件夹不存在或路径无效")
    
    def open_excel_file(self, file_path: str):
        """打开Excel文件"""
        if file_path and os.path.exists(file_path):
            url = QUrl.fromLocalFile(file_path)
            QDesktopServices.openUrl(url)
        else:
            QMessageBox.warning(self, "提示", "Excel文件不存在或路径无效")
    
    def on_error(self, error_msg):
        self.reset_ui()
        QMessageBox.critical(self, "错误", error_msg)
    
    def reset_ui(self):
        self.btn_start.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.status_label.setText("就绪")
        self.progress.setValue(0)
    
    def show_statistics(self):
        QMessageBox.information(self, "统计", "统计功能待实现")
    
    def show_history(self):
        QMessageBox.information(self, "历史记录", "历史记录功能待实现")
    
    def show_about(self):
        QMessageBox.about(
            self, "关于",
            "发票智能整理工具 V7.0 - 纯API版\n\n"
            "支持功能：\n"
            "• 增值税发票自动识别（百度API）\n"
            "• 12306铁路电子客票识别（百度API）\n"
            "• 多子表Excel导出\n"
            "• 重复发票检测\n\n"
            "说明：本版本只使用百度API进行识别，\n"
            "需要配置API密钥并联网使用。"
        )
    
    def show_print_dialog(self):
        """显示打印设置对话框"""
        from gui.print_dialog import PrintSettingsDialog
        
        # 使用保存的类型汇总
        type_stats = self.last_type_summary or {}
        
        # 获取输出文件夹
        output_folder = self.last_output_folder
        if not output_folder:
            QMessageBox.warning(self, "提示", "未找到输出文件夹")
            return
        
        # 显示打印对话框
        PrintSettingsDialog.show_print_dialog(type_stats, str(output_folder), self)
    
    # 拖拽支持
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            path = urls[0].toLocalFile()
            if Path(path).is_dir():
                self._set_input_folder(path)
            else:
                QMessageBox.warning(self, "提示", "请拖拽文件夹，不要拖拽文件")
    
    def closeEvent(self, event):
        """窗口关闭时保存配置"""
        self.config.save_config()
        event.accept()


def main():
    app = QApplication(sys.argv)
    window = MainWindowV2()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
