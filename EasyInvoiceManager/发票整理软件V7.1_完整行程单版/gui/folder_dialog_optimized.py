"""
Windows风格的文件夹选择对话框 - 异步优化版
优化点：
1. 延迟子目录检测（默认显示展开箭头，点击时再验证）
2. 异步文件列表加载，避免阻塞UI
3. 快捷方式懒解析
4. 分批加载大量文件，显示进度
"""
import os
from pathlib import Path
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTreeWidget, QTreeWidgetItem,
    QPushButton, QLabel, QLineEdit, QMessageBox, QListWidget, QListWidgetItem,
    QSplitter, QAbstractItemView, QDialogButtonBox, QWidget, QFrame,
    QToolButton, QMenu, QAction, QHeaderView, QTableWidget, QTableWidgetItem,
    QFileIconProvider, QStyledItemDelegate, QStyle, QSizePolicy, QApplication,
    QFileDialog, QProgressBar
)
from PyQt5.QtCore import Qt, QSize, pyqtSignal, QDir, QEvent, QTimer, QThread, QMutex
from PyQt5.QtGui import QIcon, QFont, QPixmap, QColor, QFontMetrics


class FileItem:
    """文件/文件夹项目 - 简化版，延迟解析快捷方式"""
    def __init__(self, path: str, is_dir: bool = False, stat_result=None):
        self.path = path
        self.name = Path(path).name if path else ''
        self.is_dir = is_dir
        self.size = 0
        self.modified = ""
        self.is_shortcut = False
        self.target_path = None
        self._stat_loaded = False
        self._stat_result = stat_result
        
    def _load_stat(self):
        """延迟加载文件信息"""
        if self._stat_loaded:
            return
        self._stat_loaded = True
        
        if not self.is_dir and self.path and Path(self.path).exists():
            try:
                if self._stat_result is not None:
                    stat = self._stat_result
                else:
                    stat = Path(self.path).stat()
                self.size = stat.st_size if hasattr(stat, 'st_size') else 0
                from datetime import datetime
                mtime = stat.st_mtime if hasattr(stat, 'st_mtime') else 0
                self.modified = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M') if mtime else ''
            except (OSError, IOError):
                pass
    
    def resolve_shortcut_lazy(self):
        """延迟解析快捷方式 - 只在需要时调用"""
        if self.is_shortcut or not self.path:
            return
            
        # 检测是否为快捷方式
        if self.path.lower().endswith('.lnk'):
            self.is_shortcut = True
            self.target_path = self._resolve_shortcut(self.path)
        elif Path(self.path).is_symlink():
            self.is_shortcut = True
            try:
                self.target_path = str(Path(self.path).resolve())
            except:
                pass
    
    def _resolve_shortcut(self, lnk_path: str) -> str:
        """解析Windows快捷方式指向的真实路径"""
        try:
            import win32com.client
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(lnk_path)
            target = shortcut.Targetpath
            if target and os.path.exists(target):
                return target
        except Exception:
            pass
        return None
    
    def get_size_str(self) -> str:
        self._load_stat()
        if self.is_dir:
            return ""
        size = self.size
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024:
                if size < 10 and unit != 'B':
                    return f"{size:.2f}{unit}"
                return f"{size:.0f}{unit}"
            size /= 1024
        return f"{size:.0f}TB"


class FileListLoader(QThread):
    """异步文件列表加载器"""
    # 信号：items(列表), 当前数量, 总数
    items_batch = pyqtSignal(list, int, int)
    loading_finished = pyqtSignal(list, list, list, list, bool)  # dirs, shortcut_dirs, invoice_files, archive_files, is_truncated
    loading_error = pyqtSignal(str)
    progress_update = pyqtSignal(int, int, str)  # current, total, message
    
    def __init__(self, path: str, max_files: int = 500):
        super().__init__()
        self.path = path
        self.max_files = max_files
        self._is_running = True
        self._mutex = QMutex()
        
        # 发票和压缩包扩展名
        self.invoice_exts = ['.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif']
        self.archive_exts = ['.zip', '.rar', '.7z', '.tar', '.gz', '.bz2', '.tgz', '.tar.gz', '.tar.bz2']
        
    def stop(self):
        """安全停止线程"""
        self._mutex.lock()
        self._is_running = False
        self._mutex.unlock()
        
    def is_running(self):
        self._mutex.lock()
        result = self._is_running
        self._mutex.unlock()
        return result
        
    def run(self):
        try:
            if not self.path or not os.path.exists(self.path):
                self.loading_error.emit("路径不存在")
                return
                
            entries = []
            entry_count = 0
            
            # 第一阶段：快速扫描所有条目
            self.progress_update.emit(0, 0, "正在扫描...")
            for entry in os.scandir(self.path):
                if not self.is_running():
                    return
                entries.append(entry)
                entry_count += 1
                if entry_count > self.max_files * 3:  # 提前截断
                    break
            
            # 第二阶段：分类处理
            dirs = []
            shortcut_dirs = []
            files = []
            
            total = len(entries)
            for i, entry in enumerate(entries):
                if not self.is_running():
                    return
                    
                if i % 50 == 0:  # 每50个更新一次进度
                    self.progress_update.emit(i, total, f"正在分析... ({i}/{total})")
                
                try:
                    # 普通文件夹
                    if entry.is_dir(follow_symlinks=False):
                        dirs.append((entry.name, entry.path, False))
                    # 符号链接
                    elif entry.is_symlink():
                        try:
                            target = Path(entry.path).resolve()
                            if target.is_dir():
                                shortcut_dirs.append((entry.name, str(target), True))
                            elif target.is_file():
                                stat = target.stat() if target.exists() else None
                                files.append((entry.name, str(target), target.suffix.lower(), True, stat))
                        except:
                            pass
                    # Windows快捷方式 - 延迟解析，先记录
                    elif entry.is_file() and entry.name.lower().endswith('.lnk'):
                        # 暂不解析，当作普通条目处理
                        try:
                            stat = entry.stat()
                        except:
                            stat = None
                        # 标记为待解析快捷方式
                        files.append((entry.name, entry.path, '.lnk', False, stat, True))
                    # 普通文件
                    elif entry.is_file():
                        try:
                            stat = entry.stat()
                        except:
                            stat = None
                        files.append((entry.name, entry.path, Path(entry.name).suffix.lower(), False, stat, False))
                except (OSError, IOError):
                    continue
            
            if not self.is_running():
                return
                
            # 第三阶段：分离发票和压缩包
            invoice_files_info = []
            archive_files_info = []
            
            for file_info in files:
                if len(file_info) == 6:
                    name, target_path, ext, is_shortcut, stat, is_lnk_pending = file_info
                else:
                    name, target_path, ext, is_shortcut, stat = file_info
                    is_lnk_pending = False
                
                # 如果是待解析的快捷方式，在这里解析
                if is_lnk_pending and self.is_running():
                    try:
                        import win32com.client
                        shell = win32com.client.Dispatch("WScript.Shell")
                        shortcut = shell.CreateShortCut(target_path)
                        target = shortcut.Targetpath
                        if target and os.path.isdir(target):
                            shortcut_dirs.append((name.replace('.lnk', ''), target, True))
                            continue
                        elif target and os.path.isfile(target):
                            ext = Path(target).suffix.lower()
                            is_shortcut = True
                            target_path = target
                            try:
                                stat = os.stat(target) if os.path.exists(target) else None
                            except:
                                stat = None
                    except:
                        pass
                
                if ext in self.invoice_exts:
                    invoice_files_info.append((name, target_path, ext, is_shortcut, stat))
                elif ext in self.archive_exts or any(target_path.lower().endswith(ae) for ae in ['.tar.gz', '.tar.bz2', '.tgz']):
                    archive_files_info.append((name, target_path, ext, is_shortcut, stat))
            
            if not self.is_running():
                return
            
            # 检查是否超出限制
            total_items = len(dirs) + len(shortcut_dirs) + len(invoice_files_info) + len(archive_files_info)
            is_truncated = False
            if total_items > self.max_files:
                available_slots = self.max_files - len(dirs) - len(shortcut_dirs) - len(archive_files_info)
                if available_slots < 0:
                    archive_files_info = archive_files_info[:self.max_files]
                    invoice_files_info = []
                else:
                    invoice_files_info = invoice_files_info[:available_slots]
                is_truncated = True
            
            # 排序
            dirs.sort(key=lambda x: x[0].lower())
            shortcut_dirs.sort(key=lambda x: x[0].lower())
            
            self.loading_finished.emit(dirs, shortcut_dirs, invoice_files_info, archive_files_info, is_truncated)
            
        except Exception as e:
            self.loading_error.emit(str(e))


class BreadcrumbButton(QToolButton):
    """路径面包屑按钮"""
    path_clicked = pyqtSignal(str)
    
    def __init__(self, text: str, path: str, parent=None):
        super().__init__(parent)
        self.full_path = path
        self.setText(text)
        self.setStyleSheet("""
            QToolButton {
                border: none;
                padding: 2px 6px;
                background-color: transparent;
                color: #0066cc;
                font-size: 14px;
            }
            QToolButton:hover {
                background-color: #e5f3ff;
                border-radius: 2px;
                text-decoration: underline;
            }
        """)
        self.setCursor(Qt.PointingHandCursor)
        self.clicked.connect(lambda: self.path_clicked.emit(self.full_path))


class FolderDialog(QDialog):
    """Windows风格的文件夹选择对话框 - 异步优化版"""
    
    folder_selected = pyqtSignal(str)
    
    ICON_FOLDER = "文件夹"
    ICON_FILE = "文件"
    ICON_PDF = "PDF"
    ICON_IMAGE = "图片"

    def __init__(self, parent=None, initial_path=None):
        super().__init__(parent)
        self.setWindowTitle("选择发票文件夹")
        self.setGeometry(100, 100, 1200, 650)
        self.setMinimumSize(900, 500)
        
        self.selected_folder = None
        self.current_path = None
        self.history = []
        self.history_index = -1
        
        self.invoice_exts = ['.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif']
        
        # 文件加载限制参数
        self.max_files_to_show = 500
        self.is_loading = False
        self.loader = None  # 文件加载线程
        
        # 延迟检测子目录的缓存
        self._checked_paths = set()  # 已检测过子目录的路径
        
        self.init_ui()
        
        if initial_path and Path(initial_path).exists():
            initial_path = str(Path(initial_path).resolve())
        else:
            initial_path = str(Path.home())
        
        self._delayed_navigate(initial_path)
    
    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(4)
        main_layout.setContentsMargins(10, 10, 10, 10)
        
        # === 顶部工具栏 ===
        toolbar = QHBoxLayout()
        toolbar.setSpacing(6)
        
        # 导航按钮组
        nav_buttons = QHBoxLayout()
        nav_buttons.setSpacing(4)
        
        self.btn_back = QPushButton("<")
        self.btn_back.setFixedSize(36, 32)
        self.btn_back.setToolTip("后退")
        self.btn_back.setEnabled(False)
        self.btn_back.clicked.connect(self.go_back)
        nav_buttons.addWidget(self.btn_back)
        
        self.btn_forward = QPushButton(">")
        self.btn_forward.setFixedSize(36, 32)
        self.btn_forward.setToolTip("前进")
        self.btn_forward.setEnabled(False)
        self.btn_forward.clicked.connect(self.go_forward)
        nav_buttons.addWidget(self.btn_forward)
        
        self.btn_up = QPushButton("^")
        self.btn_up.setFixedSize(36, 32)
        self.btn_up.setToolTip("上级目录")
        self.btn_up.clicked.connect(self.go_up)
        nav_buttons.addWidget(self.btn_up)
        
        btn_refresh = QPushButton("R")
        btn_refresh.setFixedSize(36, 32)
        btn_refresh.setToolTip("刷新")
        btn_refresh.clicked.connect(self.refresh)
        nav_buttons.addWidget(btn_refresh)
        
        toolbar.addLayout(nav_buttons)
        toolbar.addSpacing(8)
        
        # 面包屑导航
        breadcrumb_container = QWidget()
        breadcrumb_container.setStyleSheet("""
            QWidget {
                background-color: #f8f9fa;
                border: 1px solid #e0e0e0;
                border-radius: 3px;
            }
        """)
        self.breadcrumb_layout = QHBoxLayout(breadcrumb_container)
        self.breadcrumb_layout.setSpacing(0)
        self.breadcrumb_layout.setContentsMargins(6, 2, 6, 2)
        self.breadcrumb_layout.setAlignment(Qt.AlignLeft)
        toolbar.addWidget(breadcrumb_container, stretch=1)
        
        main_layout.addLayout(toolbar)
        
        # === 快速访问栏 ===
        quick_bar = QHBoxLayout()
        quick_bar.setSpacing(8)
        
        quick_label = QLabel("快速访问:")
        quick_label.setStyleSheet("font-size: 14px; color: #666;")
        quick_bar.addWidget(quick_label)
        
        quick_items = [
            ("桌面", lambda: self.navigate_to(str(Path.home() / "Desktop"), True)),
            ("此电脑", self.show_computer),
            ("文档", lambda: self.navigate_to(str(Path.home() / "Documents"), True)),
            ("下载", lambda: self.navigate_to(str(Path.home() / "Downloads"), True)),
        ]
        
        for text, callback in quick_items:
            btn = QPushButton(text)
            btn.setFixedHeight(24)
            btn.setStyleSheet("""
                QPushButton {
                    border: 1px solid #d0d0d0;
                    padding: 2px 12px;
                    background-color: #ffffff;
                    font-size: 14px;
                    border-radius: 2px;
                }
                QPushButton:hover {
                    background-color: #f0f0f0;
                    border-color: #b0b0b0;
                }
            """)
            btn.clicked.connect(callback)
            quick_bar.addWidget(btn)
        
        quick_bar.addStretch()
        main_layout.addLayout(quick_bar)
        
        # === 加载进度条（默认隐藏）===
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)  # 无限循环模式
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedHeight(3)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: none;
                background-color: #e0e0e0;
            }
            QProgressBar::chunk {
                background-color: #2196F3;
            }
        """)
        self.progress_bar.hide()
        main_layout.addWidget(self.progress_bar)
        
        # === 主内容区 ===
        main_splitter = QSplitter(Qt.Horizontal)
        main_splitter.setHandleWidth(6)
        main_splitter.setStyleSheet("""
            QSplitter::handle {
                background-color: #e0e0e0;
            }
            QSplitter::handle:hover {
                background-color: #c0c0c0;
            }
        """)
        
        # --- 左侧：文件夹树 ---
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(4)
        
        left_header = QLabel("文件夹")
        left_header.setStyleSheet("""
            font-weight: bold; 
            padding: 8px; 
            background-color: #f5f5f5; 
            border: 1px solid #ddd;
            border-bottom: none;
            font-size: 14px;
        """)
        left_layout.addWidget(left_header)
        
        self.tree_folders = QTreeWidget()
        self.tree_folders.setHeaderHidden(True)
        self.tree_folders.setColumnCount(1)
        self.tree_folders.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tree_folders.setUniformRowHeights(True)
        self.tree_folders.setAnimated(False)
        self.tree_folders.setStyleSheet("""
            QTreeWidget {
                border: 1px solid #ddd;
                font-size: 14px;
            }
            QTreeWidget::item {
                height: 22px;
                padding: 2px;
            }
            QTreeWidget::item:selected {
                background-color: #d0e8ff;
                color: #000;
            }
            QTreeWidget::item:hover {
                background-color: #e8f4ff;
            }
        """)
        self.tree_folders.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.tree_folders.itemClicked.connect(self.on_tree_item_clicked)
        self.tree_folders.itemExpanded.connect(self.on_tree_item_expanded)
        left_layout.addWidget(self.tree_folders)
        
        left_widget.setMinimumWidth(180)
        left_widget.setMaximumWidth(600)
        main_splitter.addWidget(left_widget)
        
        # --- 右侧：文件列表区 ---
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(4)
        
        # 当前路径显示
        path_frame = QFrame()
        path_frame.setStyleSheet("""
            QFrame {
                background-color: #fff8e1;
                border: 1px solid #ffe082;
                border-radius: 2px;
            }
        """)
        path_layout = QHBoxLayout(path_frame)
        path_layout.setContentsMargins(8, 4, 8, 4)
        
        path_label = QLabel("当前:")
        path_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        path_layout.addWidget(path_label)
        
        self.lbl_current_path = QLabel()
        self.lbl_current_path.setStyleSheet("font-size: 14px; color: #333;")
        self.lbl_current_path.setWordWrap(True)
        path_layout.addWidget(self.lbl_current_path, stretch=1)
        
        right_layout.addWidget(path_frame)
        
        # 文件表格
        self.table_files = QTableWidget()
        self.table_files.setColumnCount(4)
        self.table_files.setHorizontalHeaderLabels(["名称", "类型", "大小", "修改日期"])
        
        self.table_files.horizontalHeader().setStyleSheet("""
            QHeaderView::section {
                background-color: #f5f5f5;
                padding: 6px;
                border: 1px solid #ddd;
                font-weight: bold;
                font-size: 14px;
            }
        """)
        
        header = self.table_files.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        
        self.table_files.setColumnWidth(0, 300)
        self.table_files.setColumnWidth(1, 60)
        self.table_files.setColumnWidth(2, 60)
        self.table_files.setColumnWidth(3, 120)
        
        self.table_files.verticalHeader().setDefaultSectionSize(28)
        self.table_files.verticalHeader().setVisible(False)
        self.table_files.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table_files.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table_files.setAlternatingRowColors(True)
        self.table_files.setShowGrid(False)
        self.table_files.setStyleSheet("""
            QTableWidget {
                border: 1px solid #ddd;
                font-size: 14px;
            }
            QTableWidget::item {
                padding: 4px 8px;
                border-bottom: 1px solid #f0f0f0;
            }
            QTableWidget::item:selected {
                background-color: #d0e8ff;
            }
            QTableWidget::item:hover {
                background-color: #e8f4ff;
            }
        """)
        self.table_files.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.table_files.cellDoubleClicked.connect(self.on_file_double_clicked)
        right_layout.addWidget(self.table_files)
        
        # 统计信息
        self.lbl_stats = QLabel("准备就绪")
        self.lbl_stats.setStyleSheet("""
            color: #666; 
            font-size: 13px;
            padding: 6px;
            background-color: #f8f9fa;
            border: 1px solid #e0e0e0;
        """)
        right_layout.addWidget(self.lbl_stats)
        
        main_splitter.addWidget(right_widget)
        main_splitter.setSizes([300, 900])
        main_splitter.setStretchFactor(0, 0)
        main_splitter.setStretchFactor(1, 1)
        main_layout.addWidget(main_splitter, stretch=1)
        
        # === 底部操作区 ===
        bottom_frame = QFrame()
        bottom_frame.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border-top: 1px solid #e0e0e0;
            }
        """)
        bottom_layout = QHBoxLayout(bottom_frame)
        bottom_layout.setContentsMargins(10, 8, 10, 8)
        bottom_layout.setSpacing(10)
        
        selection_layout = QHBoxLayout()
        selection_label = QLabel("已选择:")
        selection_label.setStyleSheet("font-size: 14px; color: #666;")
        selection_layout.addWidget(selection_label)
        
        self.lbl_selected = QLabel("未选择")
        self.lbl_selected.setStyleSheet("""
            font-weight: bold; 
            color: #1976d2;
            font-size: 14px;
            padding: 6px 10px;
            background-color: #e3f2fd;
            border-radius: 3px;
        """)
        selection_layout.addWidget(self.lbl_selected)
        selection_layout.addStretch()
        
        bottom_layout.addLayout(selection_layout, stretch=1)
        
        self.btn_select = QPushButton("选择此文件夹")
        self.btn_select.setFixedSize(120, 32)
        self.btn_select.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                font-size: 14px;
                border-radius: 3px;
                border: none;
            }
            QPushButton:hover {
                background-color: #43a047;
            }
            QPushButton:pressed {
                background-color: #388e3c;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        self.btn_select.setEnabled(False)
        self.btn_select.clicked.connect(self.accept_folder)
        bottom_layout.addWidget(self.btn_select)
        
        btn_cancel = QPushButton("取消")
        btn_cancel.setFixedSize(80, 32)
        btn_cancel.setStyleSheet("""
            QPushButton {
                font-size: 14px;
                border-radius: 3px;
                border: 1px solid #ccc;
                background-color: #ffffff;
            }
            QPushButton:hover {
                background-color: #f5f5f5;
            }
        """)
        btn_cancel.clicked.connect(self.reject)
        bottom_layout.addWidget(btn_cancel)
        
        main_layout.addWidget(bottom_frame)
    
    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.table_files.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
    
    def closeEvent(self, event):
        """关闭时确保停止后台线程"""
        self._cancel_loading()
        event.accept()
    
    def _cancel_loading(self):
        """取消正在进行的加载"""
        if self.loader and self.loader.isRunning():
            self.loader.stop()
            self.loader.wait(500)  # 等待最多500ms
            self.loader = None
        self.is_loading = False
        self.progress_bar.hide()
    
    def _delayed_navigate(self, path: str):
        """延迟导航：打开对话框时不立即加载文件列表"""
        if not path or not os.path.exists(path):
            path = str(Path.home())
        
        path = str(Path(path).resolve())
        
        self.history = [path]
        self.history_index = 0
        
        self.update_nav_buttons()
        self.current_path = path
        
        self.update_breadcrumb(path)
        self.update_tree(path)
        
        # 清空文件列表，显示提示
        self.table_files.setRowCount(0)
        self.lbl_stats.setText("请点击左侧文件夹查看内容，或点击「选择此文件夹」确认")
        
        display_path = path
        if len(display_path) > 80:
            display_path = "..." + display_path[-77:]
        self.lbl_current_path.setText(display_path)
        self.lbl_current_path.setToolTip(path)
        
        folder_name = Path(path).name or path
        self.lbl_selected.setText(folder_name)
        self.btn_select.setEnabled(True)
    
    def navigate_to(self, path: str, add_history=True):
        """导航到指定路径"""
        if not path or path.startswith("-"):
            return
        
        # 取消之前的加载
        self._cancel_loading()
        
        # 解析快捷方式
        if path and path.lower().endswith('.lnk'):
            try:
                import win32com.client
                shell = win32com.client.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(path)
                resolved_path = shortcut.Targetpath
                if resolved_path and os.path.exists(resolved_path):
                    path = resolved_path
            except:
                pass
        
        if not path or not os.path.exists(path):
            QMessageBox.warning(self, "错误", f"路径不存在: {path}")
            return
        
        path = str(Path(path).resolve())
        
        if add_history:
            self.history = self.history[:self.history_index + 1]
            self.history.append(path)
            self.history_index = len(self.history) - 1
        
        self.update_nav_buttons()
        self.current_path = path
        
        self.update_breadcrumb(path)
        self.update_tree(path)
        self.update_file_list_async(path)  # 改为异步加载
        
        display_path = path
        if len(display_path) > 80:
            display_path = "..." + display_path[-77:]
        self.lbl_current_path.setText(display_path)
        self.lbl_current_path.setToolTip(path)
        
        folder_name = Path(path).name or path
        self.lbl_selected.setText(folder_name)
        self.btn_select.setEnabled(True)
    
    def update_file_list_async(self, path: str):
        """异步更新文件列表"""
        self.table_files.setRowCount(0)
        self.is_loading = True
        self.progress_bar.show()
        self.lbl_stats.setText("正在加载文件列表...")
        
        # 创建并启动加载线程
        self.loader = FileListLoader(path, self.max_files_to_show)
        self.loader.progress_update.connect(self._on_load_progress)
        self.loader.loading_finished.connect(self._on_load_finished)
        self.loader.loading_error.connect(self._on_load_error)
        self.loader.start()
    
    def _on_load_progress(self, current, total, message):
        """加载进度更新"""
        self.lbl_stats.setText(f"{message}")
        QApplication.processEvents()
    
    def _on_load_finished(self, dirs, shortcut_dirs, invoice_files_info, archive_files_info, is_truncated):
        """加载完成回调"""
        self.is_loading = False
        self.progress_bar.hide()
        
        # 在UI线程中更新表格
        total_items = len(dirs) + len(shortcut_dirs) + len(invoice_files_info) + len(archive_files_info)
        self.table_files.setRowCount(total_items)
        
        row = 0
        
        # 添加普通文件夹
        for name, path, _ in dirs:
            name_item = QTableWidgetItem(f"[文件夹] {name}")
            name_item.setData(Qt.UserRole, path)
            name_item.setToolTip(name)
            self.table_files.setItem(row, 0, name_item)
            
            type_item = QTableWidgetItem("文件夹")
            self.table_files.setItem(row, 1, type_item)
            
            size_item = QTableWidgetItem("--")
            self.table_files.setItem(row, 2, size_item)
            
            date_item = QTableWidgetItem("")
            self.table_files.setItem(row, 3, date_item)
            
            bg_color = QColor(232, 244, 255)
            for col in range(4):
                item = self.table_files.item(row, col)
                if item:
                    item.setBackground(bg_color)
            
            row += 1
        
        # 添加快捷方式文件夹
        for name, target_path, _ in shortcut_dirs:
            name_item = QTableWidgetItem(f"[链接] {name}")
            name_item.setData(Qt.UserRole, target_path)
            name_item.setToolTip(f"快捷方式: {name} -> {target_path}")
            name_item.setForeground(QColor(0, 100, 200))
            self.table_files.setItem(row, 0, name_item)
            
            type_item = QTableWidgetItem("快捷方式")
            self.table_files.setItem(row, 1, type_item)
            
            size_item = QTableWidgetItem("--")
            self.table_files.setItem(row, 2, size_item)
            
            date_item = QTableWidgetItem("")
            self.table_files.setItem(row, 3, date_item)
            
            bg_color = QColor(255, 248, 220)
            for col in range(4):
                item = self.table_files.item(row, col)
                if item:
                    item.setBackground(bg_color)
            
            row += 1
        
        # 添加发票文件
        for name, target_path, ext, is_shortcut, stat in invoice_files_info:
            icon = self.ICON_PDF if ext == '.pdf' else self.ICON_IMAGE
            
            prefix = "[链接] " if is_shortcut else ""
            name_item = QTableWidgetItem(f"{prefix}[{icon}] {name}")
            name_item.setData(Qt.UserRole, target_path)
            if is_shortcut:
                name_item.setToolTip(f"快捷方式: {name} -> {target_path}")
                name_item.setForeground(QColor(0, 100, 200))
            else:
                name_item.setToolTip(name)
                name_item.setForeground(QColor(0, 90, 160))
            self.table_files.setItem(row, 0, name_item)
            
            type_str = "快捷方式" if is_shortcut else (ext.upper()[1:] + "文件" if ext else "文件")
            type_item = QTableWidgetItem(type_str)
            self.table_files.setItem(row, 1, type_item)
            
            item = FileItem(target_path, stat_result=stat)
            size_item = QTableWidgetItem(item.get_size_str())
            self.table_files.setItem(row, 2, size_item)
            
            date_item = QTableWidgetItem(item.modified)
            self.table_files.setItem(row, 3, date_item)
            
            row += 1
        
        # 添加压缩包文件
        for name, target_path, ext, is_shortcut, stat in archive_files_info:
            prefix = "[链接] " if is_shortcut else ""
            name_item = QTableWidgetItem(f"{prefix}[压缩包] {name}")
            name_item.setData(Qt.UserRole, target_path)
            if is_shortcut:
                name_item.setToolTip(f"快捷方式: {name} -> {target_path}")
                name_item.setForeground(QColor(0, 100, 200))
            else:
                name_item.setToolTip(name)
                name_item.setForeground(QColor(180, 0, 180))
            self.table_files.setItem(row, 0, name_item)
            
            type_item = QTableWidgetItem("压缩包")
            self.table_files.setItem(row, 1, type_item)
            
            item = FileItem(target_path, stat_result=stat)
            size_item = QTableWidgetItem(item.get_size_str())
            self.table_files.setItem(row, 2, size_item)
            
            date_item = QTableWidgetItem(item.modified)
            self.table_files.setItem(row, 3, date_item)
            
            bg_color = QColor(245, 230, 255)
            for col in range(4):
                item = self.table_files.item(row, col)
                if item:
                    item.setBackground(bg_color)
            
            row += 1
        
        # 更新统计
        stats_text = f"共 {len(dirs) + len(shortcut_dirs)} 个文件夹"
        if len(archive_files_info) > 0:
            stats_text += f"，{len(archive_files_info)} 个压缩包"
        if len(invoice_files_info) > 0:
            stats_text += f"，{len(invoice_files_info)} 个发票文件"
        if len(shortcut_dirs) > 0:
            stats_text += f"（含 {len(shortcut_dirs)} 个快捷方式）"
        if is_truncated:
            stats_text += f" [仅显示前{self.max_files_to_show}项]"
        self.lbl_stats.setText(stats_text)
        
        self.table_files.resizeColumnsToContents()
        self.table_files.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
    
    def _on_load_error(self, error_msg):
        """加载错误回调"""
        self.is_loading = False
        self.progress_bar.hide()
        self.lbl_stats.setText(f"加载失败: {error_msg}")
    
    def update_breadcrumb(self, path: str):
        """更新面包屑导航"""
        while self.breadcrumb_layout.count():
            item = self.breadcrumb_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        parts = []
        current = Path(path) if path else Path.home()
        
        if len(str(current)) == 3 and str(current).endswith(':\\'):
            parts.append((str(current), str(current)))
        else:
            while current != current.parent:
                display_name = current.name or str(current)
                if len(display_name) > 20:
                    display_name = display_name[:17] + "..."
                parts.append((display_name, str(current)))
                current = current.parent
            if str(current) != '.':
                parts.append((str(current), str(current)))
        
        parts.reverse()
        
        for i, (name, full_path) in enumerate(parts):
            if i > 0:
                sep = QLabel(">")
                sep.setStyleSheet("color: #999; padding: 0 4px; font-size: 14px;")
                self.breadcrumb_layout.addWidget(sep)
            
            btn = BreadcrumbButton(name, full_path)
            btn.path_clicked.connect(self.on_breadcrumb_clicked)
            btn.setToolTip(full_path)
            self.breadcrumb_layout.addWidget(btn)
        
        self.breadcrumb_layout.addStretch()
    
    def on_breadcrumb_clicked(self, path: str):
        self.navigate_to(path, add_history=True)
    
    def update_tree(self, current_path: str):
        """更新文件夹树 - 简化版，延迟展开"""
        self.tree_folders.clear()
        
        if os.name == 'nt':
            import string
            # 添加桌面等特殊文件夹快捷方式
            special_folders = [
                ("桌面", str(Path.home() / "Desktop")),
                ("文档", str(Path.home() / "Documents")),
                ("下载", str(Path.home() / "Downloads")),
            ]
            
            for name, folder_path in special_folders:
                if folder_path and os.path.exists(folder_path):
                    item = QTreeWidgetItem([name])
                    item.setData(0, Qt.UserRole, folder_path)
                    item.setChildIndicatorPolicy(QTreeWidgetItem.ShowIndicator)
                    self.tree_folders.addTopLevelItem(item)
            
            # 分隔线
            separator = QTreeWidgetItem(["----------"])
            separator.setFlags(Qt.NoItemFlags)
            self.tree_folders.addTopLevelItem(separator)
            
            # 添加磁盘驱动器
            for letter in string.ascii_uppercase:
                drive = f"{letter}:\\"
                if os.path.exists(drive):
                    item = QTreeWidgetItem([f"{letter}:"])
                    item.setData(0, Qt.UserRole, drive)
                    item.setChildIndicatorPolicy(QTreeWidgetItem.ShowIndicator)
                    self.tree_folders.addTopLevelItem(item)
                    
                    if current_path and current_path.upper().startswith(drive.upper()):
                        # 延迟展开，不递归
                        pass
        else:
            item = QTreeWidgetItem(["/"])
            item.setData(0, Qt.UserRole, "/")
            self.tree_folders.addTopLevelItem(item)
    
    def expand_to_path(self, parent_item: QTreeWidgetItem, target_path: str):
        """展开树到指定路径 - 简化版"""
        parent_path = parent_item.data(0, Qt.UserRole)
        
        if parent_path and parent_path.startswith("-"):
            return
        
        if not parent_path or not target_path or not target_path.upper().startswith(parent_path.upper()):
            return
        
        self.load_tree_children(parent_item)
        parent_item.setExpanded(True)
        
        for i in range(parent_item.childCount()):
            child = parent_item.child(i)
            child_path = child.data(0, Qt.UserRole)
            
            if not child_path or child_path.startswith("-"):
                continue
            
            if target_path.upper() == child_path.upper():
                self.tree_folders.setCurrentItem(child)
                child.setSelected(True)
                return
            elif target_path.upper().startswith(child_path.upper() + os.sep):
                self.expand_to_path(child, target_path)
                return
    
    def load_tree_children(self, item: QTreeWidgetItem):
        """加载树的子项 - 延迟检测子目录"""
        if item.childCount() > 0:
            return
        
        path = item.data(0, Qt.UserRole)
        
        if not path:
            return
        
        try:
            entries = []
            for entry in os.scandir(path):
                try:
                    if entry.is_dir(follow_symlinks=False):
                        entries.append((entry.name, entry.path, False))
                    elif entry.is_symlink():
                        try:
                            target = Path(entry.path).resolve()
                            if target.is_dir():
                                entries.append((entry.name, str(target), True))
                        except:
                            pass
                    # 简化了：不在这里解析.lnk快捷方式
                except (OSError, IOError):
                    continue
            
            entries.sort(key=lambda x: x[0].lower())
            
            for name, target_path, is_shortcut in entries:
                display_name = name
                if len(name) > 25:
                    display_name = name[:22] + "..."
                
                if is_shortcut:
                    display_name = "[链接] " + display_name
                
                child = QTreeWidgetItem([display_name])
                child.setData(0, Qt.UserRole, target_path)
                child.setToolTip(0, f"{name} -> {target_path}" if is_shortcut else name)
                
                # 关键优化：默认显示展开箭头，不预先检测是否有子目录
                # 只有当用户实际展开时才去检测
                child.setChildIndicatorPolicy(QTreeWidgetItem.ShowIndicator)
                
                item.addChild(child)
        except PermissionError:
            pass
    
    def on_tree_item_expanded(self, item: QTreeWidgetItem):
        """树节点展开时加载子项"""
        self.load_tree_children(item)
        # 延迟检测：如果展开后没有子项，移除展开箭头
        if item.childCount() == 0:
            item.setChildIndicatorPolicy(QTreeWidgetItem.DontShowIndicator)
    
    def on_tree_item_clicked(self, item: QTreeWidgetItem, column: int):
        path = item.data(0, Qt.UserRole)
        if not path or path.startswith("-"):
            return
        self.navigate_to(path, add_history=True)
    
    def on_file_double_clicked(self, row: int, column: int):
        item = self.table_files.item(row, 0)
        if not item:
            return
        
        path = item.data(Qt.UserRole)
        if not path:
            return
        
        try:
            is_dir = Path(path).is_dir()
        except:
            is_dir = False
        
        if is_dir:
            self.navigate_to(path, add_history=True)
        else:
            QMessageBox.information(
                self, "提示",
                f"您双击了文件: {Path(path).name}\n\n"
                "此对话框只能选择文件夹。\n"
                "系统会自动整理文件夹内的所有发票文件。"
            )
    
    def update_nav_buttons(self):
        self.btn_back.setEnabled(self.history_index > 0)
        self.btn_forward.setEnabled(self.history_index < len(self.history) - 1)
    
    def go_back(self):
        if self.history_index > 0:
            self.history_index -= 1
            self.navigate_to(self.history[self.history_index], add_history=False)
    
    def go_forward(self):
        if self.history_index < len(self.history) - 1:
            self.history_index += 1
            self.navigate_to(self.history[self.history_index], add_history=False)
    
    def go_up(self):
        if self.current_path:
            parent = str(Path(self.current_path).parent)
            if parent != self.current_path:
                self.navigate_to(parent, add_history=True)
    
    def show_computer(self):
        """显示此电脑"""
        self.tree_folders.clear()
        
        import string
        for letter in string.ascii_uppercase:
            drive = f"{letter}:\\"
            if os.path.exists(drive):
                item = QTreeWidgetItem([f"{letter}:"])
                item.setData(0, Qt.UserRole, drive)
                item.setChildIndicatorPolicy(QTreeWidgetItem.ShowIndicator)
                self.tree_folders.addTopLevelItem(item)
        
        self.current_path = ""
        self.update_breadcrumb("此电脑")
        self.table_files.setRowCount(0)
        self.lbl_stats.setText("请选择驱动器")
        self.lbl_current_path.setText("此电脑")
        self.btn_select.setEnabled(False)
        self.lbl_selected.setText("未选择")
    
    def refresh(self):
        if self.current_path:
            self.navigate_to(self.current_path, add_history=False)
    
    def accept_folder(self):
        if not self.current_path:
            QMessageBox.warning(self, "提示", "请先选择一个文件夹")
            return
        
        if Path(self.current_path).is_file():
            QMessageBox.warning(self, "错误", "不能选择文件")
            return
        
        self.selected_folder = self.current_path
        self._cancel_loading()  # 确保停止后台线程
        self.accept()
    
    def get_selected_folder(self) -> str:
        return self.selected_folder


class SimpleFolderDialog:
    """兼容层"""
    def __init__(self, parent=None, initial_path=None):
        self.parent = parent
        self.initial_path = initial_path
        self.selected_folder = None
    
    def exec_(self):
        from PyQt5.QtWidgets import QFileDialog
        folder = QFileDialog.getExistingDirectory(
            self.parent, "选择发票文件夹",
            self.initial_path or str(Path.home())
        )
        if folder:
            self.selected_folder = folder
            return QDialog.Accepted
        return QDialog.Rejected
    
    def get_selected_folder(self) -> str:
        return self.selected_folder
