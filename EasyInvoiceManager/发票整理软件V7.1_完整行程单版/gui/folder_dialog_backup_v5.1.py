"""
Windows风格的文件夹选择对话框 - 自适应布局版
"""
import os
from pathlib import Path
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTreeWidget, QTreeWidgetItem,
    QPushButton, QLabel, QLineEdit, QMessageBox, QListWidget, QListWidgetItem,
    QSplitter, QAbstractItemView, QDialogButtonBox, QWidget, QFrame,
    QToolButton, QMenu, QAction, QHeaderView, QTableWidget, QTableWidgetItem,
    QFileIconProvider, QStyledItemDelegate, QStyle, QSizePolicy, QApplication,
    QFileDialog
)
from PyQt5.QtCore import Qt, QSize, pyqtSignal, QDir, QEvent, QTimer
from PyQt5.QtGui import QIcon, QFont, QPixmap, QColor, QFontMetrics


class FileItem:
    """文件/文件夹项目"""
    def __init__(self, path: str, is_dir: bool = False, stat_result=None):
        self.path = path
        self.name = Path(path).name if path else ''
        self.is_dir = is_dir
        self.size = 0
        self.modified = ""
        self.is_shortcut = False
        self.target_path = None  # 快捷方式指向的真实路径
        
        if not is_dir and path and Path(path).exists():
            try:
                # 如果提供了stat结果，直接使用
                if stat_result is not None:
                    stat = stat_result
                else:
                    stat = Path(path).stat()
                self.size = stat.st_size if hasattr(stat, 'st_size') else 0
                from datetime import datetime
                mtime = stat.st_mtime if hasattr(stat, 'st_mtime') else 0
                self.modified = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M') if mtime else ''
            except (OSError, IOError):
                self.size = 0
                self.modified = ''
            
            # 检测是否为快捷方式
            if path and path.lower().endswith('.lnk'):
                self.is_shortcut = True
                self.target_path = self._resolve_shortcut(path)
            # 检测是否为符号链接
            elif Path(path).is_symlink():
                self.is_shortcut = True
                self.target_path = str(Path(path).resolve())
    
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
        except ImportError:
            # pywin32 未安装
            pass
        
        # 备选：尝试读取快捷方式文件（简化解析）
        try:
            with open(lnk_path, 'rb') as f:
                content = f.read()
                # 尝试找到路径字符串
                try:
                    # 查找常见的路径模式
                    idx = content.find(b'\\')
                    if idx > 0:
                        # 尝试解码
                        for i in range(max(0, idx-100), idx):
                            try:
                                path_part = content[i:idx+100].decode('utf-16-le', errors='ignore')
                                if ':' in path_part and os.path.exists(path_part.split('\x00')[0]):
                                    return path_part.split('\x00')[0]
                            except:
                                continue
                except:
                    pass
        except:
            pass
        
        return None
    
    def get_size_str(self) -> str:
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
    """Windows风格的文件夹选择对话框 - 自适应布局版"""
    
    folder_selected = pyqtSignal(str)
    
    ICON_FOLDER = "文件夹"
    ICON_FILE = "文件"
    ICON_PDF = "PDF"
    ICON_IMAGE = "图片"
    
    def __init__(self, parent=None, initial_path=None):
        super().__init__(parent)
        self.setWindowTitle("选择发票文件夹")
        self.setGeometry(100, 100, 1200, 650)  # 降低窗口高度
        # 设置最小尺寸，允许更大幅度的拉伸
        self.setMinimumSize(900, 500)
        
        self.selected_folder = None
        self.current_path = None
        self.history = []
        self.history_index = -1
        
        self.invoice_exts = ['.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif']
        
        # 文件加载限制参数
        self.max_files_to_show = 500  # 最大显示文件数量
        self.is_loading = False  # 加载状态标志
        
        self.init_ui()
        
        # 优化：如果初始路径是文件夹，只导航到父目录，避免立即加载大量文件
        if initial_path and Path(initial_path).exists():
            initial_path = str(Path(initial_path).resolve())
        else:
            initial_path = str(Path.home())
        
        # 使用延迟加载：先导航到路径，但不加载文件列表
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
        
        # 面包屑导航 - 使用可滚动的widget
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
        
        # === 主内容区 - 使用分割器 ===
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
        # 设置树形控件大小策略，使其能够自适应缩放
        self.tree_folders.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.tree_folders.itemClicked.connect(self.on_tree_item_clicked)
        self.tree_folders.itemExpanded.connect(self.on_tree_item_expanded)
        left_layout.addWidget(self.tree_folders)
        
        left_widget.setMinimumWidth(180)
        left_widget.setMaximumWidth(600)  # 放宽最大宽度限制，允许更大的左右拉伸幅度
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
        
        # 文件表格 - 使用自适应列宽
        self.table_files = QTableWidget()
        self.table_files.setColumnCount(4)
        self.table_files.setHorizontalHeaderLabels(["名称", "类型", "大小", "修改日期"])
        
        # 设置表头样式
        self.table_files.horizontalHeader().setStyleSheet("""
            QHeaderView::section {
                background-color: #f5f5f5;
                padding: 6px;
                border: 1px solid #ddd;
                font-weight: bold;
                font-size: 14px;
            }
        """)
        
        # 自适应列宽策略
        header = self.table_files.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)  # 名称列自适应
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)  # 类型列根据内容
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)  # 大小列根据内容
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # 日期列根据内容
        
        # 设置最小列宽确保表头显示完整
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
        # 设置表格大小策略，使其能够自适应缩放
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
        
        # 设置分割比例
        main_splitter.setSizes([300, 900])
        # 允许分割器在拉伸时自动调整
        main_splitter.setStretchFactor(0, 0)  # 左侧不拉伸
        main_splitter.setStretchFactor(1, 1)  # 右侧拉伸
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
        
        # 左侧：已选择显示
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
        
        # 右侧：按钮
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
        """窗口大小改变时调整列宽"""
        super().resizeEvent(event)
        # 确保名称列占据主要空间
        self.table_files.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
    
    def _delayed_navigate(self, path: str):
        """延迟导航：打开对话框时不立即加载文件列表，等待用户点击"""
        if not path or not os.path.exists(path):
            path = str(Path.home())
        
        path = str(Path(path).resolve())
        
        # 只更新基本信息，不加载文件列表
        self.history = [path]
        self.history_index = 0
        
        self.update_nav_buttons()
        self.current_path = path
        
        self.update_breadcrumb(path)
        self.update_tree(path)
        # 不调用 update_file_list，等待用户点击树节点
        
        # 清空文件列表，显示提示
        self.table_files.setRowCount(0)
        self.lbl_stats.setText("请点击左侧文件夹查看内容，或点击「选择此文件夹」确认")
        
        # 更新当前路径显示
        display_path = path
        if len(display_path) > 80:
            display_path = "..." + display_path[-77:]
        self.lbl_current_path.setText(display_path)
        self.lbl_current_path.setToolTip(path)
        
        folder_name = Path(path).name or path
        self.lbl_selected.setText(folder_name)
        self.btn_select.setEnabled(True)
    
    def navigate_to(self, path: str, add_history=True):
        """导航到指定路径（支持快捷方式）"""
        if not path or path.startswith("-"):
            return
        
        # 解析快捷方式
        if path and path.lower().endswith('.lnk'):
            try:
                import win32com.client
                shell = win32com.client.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(path)
                resolved_path = shortcut.Targetpath
                if resolved_path and os.path.exists(resolved_path):
                    path = resolved_path
            except ImportError:
                # pywin32 未安装，跳过快捷方式解析
                pass
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
        self.update_file_list(path)
        
        # 更新当前路径显示
        display_path = path
        if len(display_path) > 80:
            display_path = "..." + display_path[-77:]
        self.lbl_current_path.setText(display_path)
        self.lbl_current_path.setToolTip(path)
        
        folder_name = Path(path).name or path
        self.lbl_selected.setText(folder_name)
        self.btn_select.setEnabled(True)
    
    def update_breadcrumb(self, path: str):
        """更新面包屑导航"""
        # 清空
        while self.breadcrumb_layout.count():
            item = self.breadcrumb_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        parts = []
        current = Path(path) if path else Path.home()
        
        # 处理驱动器
        if len(str(current)) == 3 and str(current).endswith(':\\'):
            parts.append((str(current), str(current)))
        else:
            while current != current.parent:
                display_name = current.name or str(current)
                # 限制显示长度
                if len(display_name) > 20:
                    display_name = display_name[:17] + "..."
                parts.append((display_name, str(current)))
                current = current.parent
            if str(current) != '.':
                parts.append((str(current), str(current)))
        
        parts.reverse()
        
        # 创建面包屑按钮
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
        """更新文件夹树（支持桌面等特殊文件夹）"""
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
                    
                    if current_path and current_path.upper().startswith(folder_path.upper()):
                        self.expand_to_path(item, current_path)
            
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
                        self.expand_to_path(item, current_path)
        else:
            item = QTreeWidgetItem(["/"])
            item.setData(0, Qt.UserRole, "/")
            self.tree_folders.addTopLevelItem(item)
    
    def expand_to_path(self, parent_item: QTreeWidgetItem, target_path: str):
        """展开树到指定路径（支持快捷方式）"""
        parent_path = parent_item.data(0, Qt.UserRole)
        
        # 跳过分隔线
        if parent_path and parent_path.startswith("-"):
            return
        
        if not parent_path or not target_path or not target_path.upper().startswith(parent_path.upper()):
            return
        
        self.load_tree_children(parent_item)
        parent_item.setExpanded(True)
        
        for i in range(parent_item.childCount()):
            child = parent_item.child(i)
            child_path = child.data(0, Qt.UserRole)
            
            # 跳过分隔线
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
        """加载树的子项（支持快捷方式）- 使用os.scandir优化性能"""
        if item.childCount() > 0:
            return
        
        path = item.data(0, Qt.UserRole)
        
        if not path:
            return
        
        try:
            # 使用os.scandir()替代Path.iterdir()，性能更好
            entries = []
            for entry in os.scandir(path):
                try:
                    # 使用entry的stat信息，避免额外的系统调用
                    if entry.is_dir(follow_symlinks=False):
                        # 普通文件夹（不是符号链接）
                        entries.append((entry.name, entry.path, False))
                    elif entry.is_symlink():
                        # 符号链接，需要检查指向的目标
                        try:
                            target = Path(entry.path).resolve()
                            if target.is_dir():
                                entries.append((entry.name, str(target), True))
                        except:
                            pass
                    elif entry.is_file() and entry.name.lower().endswith('.lnk'):
                        # Windows快捷方式(.lnk文件)
                        try:
                            import win32com.client
                            shell = win32com.client.Dispatch("WScript.Shell")
                            shortcut = shell.CreateShortCut(entry.path)
                            target = shortcut.Targetpath
                            if target and os.path.isdir(target):
                                entries.append((entry.name.replace('.lnk', ''), target, True))
                        except ImportError:
                            # pywin32 未安装，跳过
                            pass
                        except:
                            pass
                except (OSError, IOError):
                    # 跳过无法访问的条目
                    continue
            
            # 排序
            entries.sort(key=lambda x: x[0].lower())
            
            for name, target_path, is_shortcut in entries:
                # 限制显示长度
                display_name = name
                if len(name) > 25:
                    display_name = name[:22] + "..."
                
                # 添加快捷方式标记
                if is_shortcut:
                    display_name = "[链接] " + display_name
                
                child = QTreeWidgetItem([display_name])
                child.setData(0, Qt.UserRole, target_path)
                child.setToolTip(0, f"{name} -> {target_path}" if is_shortcut else name)
                
                # 检查是否有子目录 - 同样使用os.scandir优化
                try:
                    has_subdirs = False
                    for subentry in os.scandir(target_path):
                        try:
                            if subentry.is_dir(follow_symlinks=False) or subentry.is_symlink():
                                has_subdirs = True
                                break
                        except (OSError, IOError):
                            continue
                    if has_subdirs:
                        child.setChildIndicatorPolicy(QTreeWidgetItem.ShowIndicator)
                except:
                    pass
                
                item.addChild(child)
        except PermissionError:
            pass
    
    def on_tree_item_expanded(self, item: QTreeWidgetItem):
        self.load_tree_children(item)
    
    def on_tree_item_clicked(self, item: QTreeWidgetItem, column: int):
        path = item.data(0, Qt.UserRole)
        # 跳过分隔线
        if not path or path.startswith("-"):
            return
        # 用户点击树节点时加载文件列表
        self.navigate_to(path, add_history=True)
    
    def update_file_list(self, path: str):
        """更新文件列表（支持快捷方式）- 使用os.scandir优化性能，减少文件系统调用"""
        # 如果正在加载，取消之前的加载
        if self.is_loading:
            return
        
        self.is_loading = True
        self.table_files.setRowCount(0)
        
        # 显示加载中提示
        self.lbl_stats.setText("正在加载文件列表...")
        QApplication.processEvents()
        
        if not path:
            self.is_loading = False
            return
        
        try:
            # 使用os.scandir()替代Path.iterdir()，性能更好
            # os.scandir()一次性返回包含文件类型信息的DirEntry对象
            entries = []
            entry_count = 0
            for entry in os.scandir(path):
                entries.append(entry)
                entry_count += 1
                # 如果文件太多，提前终止
                if entry_count > self.max_files_to_show * 2:
                    break
            
            # 分离文件夹和文件（包括快捷方式）
            dirs = []
            shortcut_dirs = []  # 快捷方式指向的文件夹
            files = []
            
            for entry in entries:
                try:
                    # 普通文件夹（使用is_dir(follow_symlinks=False)避免解析符号链接）
                    if entry.is_dir(follow_symlinks=False):
                        dirs.append((entry.name, entry.path, False))
                    # 符号链接
                    elif entry.is_symlink():
                        try:
                            target = Path(entry.path).resolve()
                            if target.is_dir():
                                shortcut_dirs.append((entry.name, str(target), True))
                            elif target.is_file():
                                files.append((entry.name, str(target), target.suffix.lower(), True, target.stat()))
                        except:
                            pass
                    # Windows快捷方式(.lnk)
                    elif entry.is_file() and entry.name.lower().endswith('.lnk'):
                        try:
                            import win32com.client
                            shell = win32com.client.Dispatch("WScript.Shell")
                            shortcut = shell.CreateShortCut(entry.path)
                            target = shortcut.Targetpath
                            if target and os.path.isdir(target):
                                shortcut_dirs.append((entry.name.replace('.lnk', ''), target, True))
                            elif target and os.path.isfile(target):
                                ext = Path(target).suffix.lower()
                                stat = os.stat(target) if os.path.exists(target) else None
                                files.append((entry.name.replace('.lnk', ''), target, ext, True, stat))
                        except ImportError:
                            # pywin32 未安装，跳过
                            pass
                        except:
                            pass
                    # 普通文件
                    elif entry.is_file():
                        # 使用entry.stat()避免额外的系统调用
                        try:
                            stat = entry.stat()
                        except:
                            stat = None
                        files.append((entry.name, entry.path, Path(entry.name).suffix.lower(), False, stat))
                except (OSError, IOError):
                    # 跳过无法访问的条目
                    continue
            
            # 排序
            dirs.sort(key=lambda x: x[0].lower())
            shortcut_dirs.sort(key=lambda x: x[0].lower())
            
            # 分离发票文件和压缩包文件
            invoice_files_info = []
            archive_files_info = []
            
            # 压缩包扩展名
            archive_exts = ['.zip', '.rar', '.7z', '.tar', '.gz', '.bz2', '.tgz', '.tar.gz', '.tar.bz2']
            
            for name, target_path, ext, is_shortcut, stat in files:
                if ext in self.invoice_exts:
                    invoice_files_info.append((name, target_path, ext, is_shortcut, stat))
                elif ext in archive_exts or any(target_path.lower().endswith(ae) for ae in ['.tar.gz', '.tar.bz2', '.tgz']):
                    archive_files_info.append((name, target_path, ext, is_shortcut, stat))
            
            # 检查是否超出限制
            total_items = len(dirs) + len(shortcut_dirs) + len(invoice_files_info) + len(archive_files_info)
            is_truncated = False
            if total_items > self.max_files_to_show:
                # 优先保留文件夹和压缩包，然后限制文件数量
                available_slots = self.max_files_to_show - len(dirs) - len(shortcut_dirs) - len(archive_files_info)
                if available_slots < 0:
                    # 如果文件夹+压缩包就超限制了，只保留压缩包，限制文件夹
                    archive_files_info = archive_files_info[:self.max_files_to_show]
                    invoice_files_info = []
                else:
                    invoice_files_info = invoice_files_info[:available_slots]
                is_truncated = True
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
                
                # 设置文件夹背景色
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
                
                # 设置快捷方式背景色
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
                
                # 使用预获取的stat信息创建FileItem
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
                    name_item.setForeground(QColor(180, 0, 180))  # 紫色标记压缩包
                self.table_files.setItem(row, 0, name_item)
                
                type_item = QTableWidgetItem("压缩包")
                self.table_files.setItem(row, 1, type_item)
                
                # 使用预获取的stat信息创建FileItem
                item = FileItem(target_path, stat_result=stat)
                size_item = QTableWidgetItem(item.get_size_str())
                self.table_files.setItem(row, 2, size_item)
                
                date_item = QTableWidgetItem(item.modified)
                self.table_files.setItem(row, 3, date_item)
                
                # 设置压缩包背景色（浅紫色）
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
                stats_text += f" [仅显示前{self.max_files_to_show}项，文件夹内文件过多]"
            self.lbl_stats.setText(stats_text)
            
            # 调整列宽以适应内容
            self.table_files.resizeColumnsToContents()
            # 但名称列保持自适应
            self.table_files.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
            
        except PermissionError:
            self.lbl_stats.setText("无法访问此文件夹")
        finally:
            self.is_loading = False
    
    def on_file_double_clicked(self, row: int, column: int):
        item = self.table_files.item(row, 0)
        if not item:
            return
        
        path = item.data(Qt.UserRole)
        if not path:
            return
        
        # 检查路径是否为文件夹（包括快捷方式指向的文件夹）
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
