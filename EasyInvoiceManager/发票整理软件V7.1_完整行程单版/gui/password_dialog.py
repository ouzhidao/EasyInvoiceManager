"""
压缩包密码输入对话框
"""
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
    QLineEdit, QPushButton, QCheckBox, QMessageBox
)
from PyQt5.QtCore import Qt


class ArchivePasswordDialog(QDialog):
    """压缩包密码输入对话框"""
    
    # 类变量：记住的密码
    _remembered_passwords = {}
    _skip_archives = set()
    
    def __init__(self, archive_name: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("压缩包需要密码")
        self.setFixedWidth(400)
        self.archive_name = archive_name
        self.password = None
        self.skip = False
        self.remember = False
        
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # 提示信息
        info_label = QLabel(f"压缩包需要密码才能解压：\n\n📦 {self.archive_name}")
        info_label.setWordWrap(True)
        info_label.setStyleSheet("font-size: 14px; padding: 10px; background-color: #fff3cd; border-radius: 5px;")
        layout.addWidget(info_label)
        
        # 密码输入
        password_layout = QHBoxLayout()
        password_layout.addWidget(QLabel("密码："))
        
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.setPlaceholderText("请输入压缩包密码")
        self.password_input.setStyleSheet("padding: 8px; font-size: 14px;")
        password_layout.addWidget(self.password_input)
        
        # 显示密码复选框
        self.show_password = QCheckBox("显示密码")
        self.show_password.stateChanged.connect(self.toggle_password_visibility)
        password_layout.addWidget(self.show_password)
        
        layout.addLayout(password_layout)
        
        # 记住密码选项
        self.remember_checkbox = QCheckBox("记住密码（本次任务有效）")
        self.remember_checkbox.setStyleSheet("color: #666;")
        layout.addWidget(self.remember_checkbox)
        
        # 按钮
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        self.skip_btn = QPushButton("⏭ 跳过此压缩包")
        self.skip_btn.setStyleSheet("""
            QPushButton {
                padding: 8px 16px;
                font-size: 13px;
                background-color: #ffc107;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #e0a800;
            }
        """)
        self.skip_btn.clicked.connect(self.on_skip)
        button_layout.addWidget(self.skip_btn)
        
        button_layout.addSpacing(10)
        
        cancel_btn = QPushButton("取消")
        cancel_btn.setStyleSheet("padding: 8px 16px; font-size: 13px;")
        cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(cancel_btn)
        
        ok_btn = QPushButton("✓ 确定")
        ok_btn.setStyleSheet("""
            QPushButton {
                padding: 8px 16px;
                font-size: 13px;
                background-color: #28a745;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #218838;
            }
        """)
        ok_btn.clicked.connect(self.on_ok)
        ok_btn.setDefault(True)
        button_layout.addWidget(ok_btn)
        
        layout.addLayout(button_layout)
        
        # 设置焦点
        self.password_input.setFocus()
    
    def toggle_password_visibility(self, state):
        """切换密码显示/隐藏"""
        if state == Qt.Checked:
            self.password_input.setEchoMode(QLineEdit.Normal)
        else:
            self.password_input.setEchoMode(QLineEdit.Password)
    
    def on_skip(self):
        """跳过此压缩包"""
        self.skip = True
        # 添加到跳过集合
        ArchivePasswordDialog._skip_archives.add(self.archive_name)
        self.accept()
    
    def on_ok(self):
        """确定"""
        password = self.password_input.text().strip()
        if not password:
            QMessageBox.warning(self, "提示", "请输入密码")
            return
        
        self.password = password
        self.remember = self.remember_checkbox.isChecked()
        
        # 如果记住密码，保存到类变量
        if self.remember:
            ArchivePasswordDialog._remembered_passwords[self.archive_name] = password
        
        self.accept()
    
    @classmethod
    def get_password(cls, archive_name: str, parent=None) -> tuple:
        """
        获取压缩包密码
        返回: (password, skip)
        password: 密码字符串，None表示取消
        skip: True表示用户选择跳过
        """
        # 检查是否已跳过
        if archive_name in cls._skip_archives:
            return None, True
        
        # 检查是否有记住的密码
        if archive_name in cls._remembered_passwords:
            return cls._remembered_passwords[archive_name], False
        
        # 显示对话框
        dialog = cls(archive_name, parent)
        result = dialog.exec_()
        
        if result == QDialog.Rejected and not dialog.skip:
            # 用户点击取消（不是跳过）
            return None, False
        
        return dialog.password, dialog.skip
    
    @classmethod
    def reset_session(cls):
        """重置会话（新任务开始时调用）"""
        cls._remembered_passwords.clear()
        cls._skip_archives.clear()
