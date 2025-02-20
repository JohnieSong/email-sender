from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                           QPushButton)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
from styles import MODERN_STYLE
import sys
import os

def get_resource_path(relative_path):
    """获取资源文件的绝对路径"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle('关于')
        self.setStyleSheet(MODERN_STYLE)
        self.setFixedSize(400, 300)  # 减小窗口尺寸
        self.setWindowFlags(Qt.Dialog | Qt.MSWindowsFixedSizeDialogHint)
        
        # 计算对话框相对于主窗口的居中位置
        if parent:
            x = parent.x() + (parent.width() - self.width()) // 2
            y = parent.y() + (parent.height() - self.height()) // 2
            self.move(x, y)
            
        self.initUI()
        
    def initUI(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # 顶部布局：logo和标题
        top_layout = QHBoxLayout()
        
        # Logo
        icon_label = QLabel()
        icon_path = get_resource_path('imgs/logo.ico')
        icon_label.setPixmap(QIcon(icon_path).pixmap(48, 48))
        top_layout.addWidget(icon_label)
        
        # 标题和版本
        title_version_layout = QVBoxLayout()
        title_label = QLabel("邮件批量发送程序")
        title_label.setStyleSheet("""
            QLabel {
                color: #40a9ff;
                font-size: 18px;
                font-weight: bold;
            }
        """)
        version_label = QLabel("版本 1.2.5")
        version_label.setStyleSheet("color: #666666;")
        
        title_version_layout.addWidget(title_label)
        title_version_layout.addWidget(version_label)
        top_layout.addLayout(title_version_layout)
        top_layout.addStretch()
        
        # 简介
        desc_label = QLabel("""
            <p style='line-height: 1.5; color: #333333;'>
            专业的批量邮件发送工具，支持多种邮箱服务商，提供模板管理和变量替换功能。
            </p>
        """)
        desc_label.setWordWrap(True)
        
        # 主要功能
        features_label = QLabel("""
            <p style='margin: 0; line-height: 1.6; color: #333333;'>
            • 支持多种邮箱服务商<br>
            • HTML邮件模板管理<br>
            • Excel导入收件人信息<br>
            • 多线程发送和状态监控
            </p>
        """)
        
        # 底部信息
        footer_label = QLabel("""
            <div style='color: #666666; font-size: 12px; text-align: center;'>
                <p>技术支持：support@bbrhub.com</p>
                <p>© 2025 BBRHub All Rights Reserved.</p>
            </div>
        """)
        footer_label.setAlignment(Qt.AlignCenter)
        
        # 添加所有组件到主布局
        layout.addLayout(top_layout)
        layout.addWidget(desc_label)
        layout.addWidget(features_label)
        layout.addStretch()
        layout.addWidget(footer_label)