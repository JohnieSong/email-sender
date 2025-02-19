from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                            QLineEdit, QPushButton, QListWidget, QMessageBox, QCheckBox)
from PyQt5.QtCore import Qt, pyqtSignal
from message_box import MessageBox
from styles import MODERN_STYLE
from database import Database
from config import Config
from logger import Logger

class ServerDialog(QDialog):
    server_updated = pyqtSignal()
    _smtp_servers_loaded = False  # 添加静态标志，记录是否已加载过服务器配置
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.config = Config()
        self.db = Database(self.config)
        self.logger = Logger()
        
        self.setWindowTitle('服务器管理')
        self.setStyleSheet(MODERN_STYLE)
        self.setMinimumSize(200, 200)
        
        # 添加默认服务器列表
        self.default_servers = [
            'QQ邮箱',
            'QQ企业邮箱',
            '163邮箱',
            '126邮箱',
            'Gmail',
            '阿里企业邮箱',
            '阿里邮箱',
            '电信邮箱',
            '搜狐邮箱',
            '新浪邮箱',
            '移动邮箱',
            '苹果邮箱',
            'Outlook',
            'AOL邮箱',
            'Yandex邮箱'
        ]
        
        self.initUI()
        
        # 只在第一次打开时加载服务器列表
        if not ServerDialog._smtp_servers_loaded:
            self.load_server_list()
            ServerDialog._smtp_servers_loaded = True
            # 记录日志，但只在第一次加载时记录
            smtp_servers = self.db.get_smtp_servers()
            self.logger.debug(f"加载到的SMTP服务器: {smtp_servers}")
        else:
            # 如果已经加载过，只更新列表显示
            self.update_server_list()
    
    def initUI(self):
        layout = QVBoxLayout()
        layout.setSpacing(10)
        
        # 服务器列表
        self.server_list = QListWidget()
        
        # 启用工具提示
        self.server_list.setMouseTracking(True)
        self.server_list.setToolTipDuration(5000)
        
        # 按钮组
        btn_layout = QHBoxLayout()
        
        # 创建按钮
        add_btn = QPushButton('添加')
        edit_btn = QPushButton('修改')
        delete_btn = QPushButton('删除')
        
        # 连接按钮事件
        add_btn.clicked.connect(self.add_server)
        edit_btn.clicked.connect(self.edit_server)
        delete_btn.clicked.connect(self.delete_server)
        
        # 将按钮添加到布局中，并设置居中对齐
        btn_layout.addStretch()  # 左侧弹性空间
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(edit_btn)
        btn_layout.addWidget(delete_btn)
        btn_layout.addStretch()  # 右侧弹性空间
        
        # 添加组件到主布局
        layout.addWidget(QLabel('发件人列表:'))
        layout.addWidget(self.server_list)
        layout.addLayout(btn_layout)
        
        self.setLayout(layout)
    
    def load_server_list(self):
        """首次加载服务器列表，包括初始化默认服务器"""
        try:
            # 检查数据库中是否已有服务器配置
            smtp_servers = self.db.get_smtp_servers()
            if not smtp_servers:
                # 如果数据库为空，初始化默认服务器配置
                self.init_default_servers()
            
            # 更新列表显示
            self.update_server_list()
            
        except Exception as e:
            self.logger.error(f"加载服务器列表失败: {str(e)}")
            MessageBox.show('错误', f'加载服务器列表失败: {str(e)}', 'error', parent=self)
    
    def update_server_list(self):
        """更新服务器列表显示"""
        try:
            self.server_list.clear()
            smtp_servers = self.db.get_smtp_servers()
            for server_type in smtp_servers.keys():
                self.server_list.addItem(server_type)
        except Exception as e:
            self.logger.error(f"更新服务器列表失败: {str(e)}")
            MessageBox.show('错误', f'更新服务器列表失败: {str(e)}', 'error', parent=self)
    
    def init_default_servers(self):
        """初始化默认服务器配置"""
        try:
            # 默认服务器配置
            default_configs = {
                'QQ邮箱': {'smtp_server': 'smtp.qq.com', 'smtp_port': 465, 'use_ssl': True, 'use_tls': False},
                'QQ企业邮箱': {'smtp_server': 'smtp.exmail.qq.com', 'smtp_port': 465, 'use_ssl': True, 'use_tls': False},
                '163邮箱': {'smtp_server': 'smtp.163.com', 'smtp_port': 465, 'use_ssl': True, 'use_tls': False},
                '126邮箱': {'smtp_server': 'smtp.126.com', 'smtp_port': 465, 'use_ssl': True, 'use_tls': False},
                'Gmail': {'smtp_server': 'smtp.gmail.com', 'smtp_port': 587, 'use_ssl': False, 'use_tls': True},
                '阿里企业邮箱': {'smtp_server': 'smtp.mxhichina.com', 'smtp_port': 465, 'use_ssl': True, 'use_tls': False},
                '阿里邮箱': {'smtp_server': 'smtp.aliyun.com', 'smtp_port': 465, 'use_ssl': True, 'use_tls': False},
                '电信邮箱': {'smtp_server': 'smtp.189.cn', 'smtp_port': 465, 'use_ssl': True, 'use_tls': False},
                '搜狐邮箱': {'smtp_server': 'smtp.sohu.com', 'smtp_port': 465, 'use_ssl': True, 'use_tls': False},
                '新浪邮箱': {'smtp_server': 'smtp.sina.com', 'smtp_port': 465, 'use_ssl': True, 'use_tls': False},
                '移动邮箱': {'smtp_server': 'smtp.139.com', 'smtp_port': 465, 'use_ssl': True, 'use_tls': False},
                '苹果邮箱': {'smtp_server': 'smtp.mail.me.com', 'smtp_port': 587, 'use_ssl': False, 'use_tls': True},
                'Outlook': {'smtp_server': 'smtp.office365.com', 'smtp_port': 587, 'use_ssl': False, 'use_tls': True},
                'AOL邮箱': {'smtp_server': 'smtp.aol.com', 'smtp_port': 587, 'use_ssl': False, 'use_tls': True},
                'Yandex邮箱': {'smtp_server': 'smtp.yandex.com', 'smtp_port': 465, 'use_ssl': True, 'use_tls': False}
            }
            
            # 批量添加默认服务器
            for server_type, config in default_configs.items():
                self.db.add_smtp_server(server_type, config)
                
            self.logger.info("已初始化默认SMTP服务器配置")
            
        except Exception as e:
            self.logger.error(f"初始化默认服务器失败: {str(e)}")
            MessageBox.show('错误', f'初始化默认服务器失败: {str(e)}', 'error', parent=self)
    
    def add_server(self):
        """添加新服务器"""
        dialog = ServerInputDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            data = dialog.get_data()
            if not all([data['server_type'], data['config']['smtp_server'], 
                       data['config']['smtp_port']]):
                MessageBox.show('错误', '请填写所有必填字段', 'warning', parent=self)
                return
            
            try:
                self.db.add_smtp_server(data['server_type'], data['config'])
                self.load_server_list()
                self.server_updated.emit()
                MessageBox.show('成功', '服务器已添加', 'info', parent=self)
            except Exception as e:
                MessageBox.show('错误', f'添加失败：{str(e)}', 'critical', parent=self)
        
    def edit_server(self):
        """编辑服务器"""
        current_item = self.server_list.currentItem()
        if not current_item:
            MessageBox.show('提示', '请先选择要修改的服务器', 'warning', parent=self)
            return
        
        server_type = current_item.text()
        server_config = self.db.get_smtp_server(server_type)
        
        # 检查是否是默认服务器
        is_default = server_type in self.default_servers
        
        dialog = ServerInputDialog(self, server_type, server_config, readonly_type=is_default)
        if dialog.exec_() == QDialog.Accepted:
            data = dialog.get_data()
            if not all([data['server_type'], data['config']['smtp_server'], 
                       data['config']['smtp_port']]):
                MessageBox.show('错误', '请填写所有必填字段', 'warning', parent=self)
                return
            
            try:
                # 如果是默认服务器，不允许修改类型
                if is_default and data['server_type'] != server_type:
                    MessageBox.show('错误', '默认服务器不能修改类型', 'warning', parent=self)
                    return
                
                # 如果服务器类型改变了，需要检查新类型是否已存在
                if data['server_type'] != server_type:
                    existing_server = self.db.get_smtp_server(data['server_type'])
                    if existing_server:
                        MessageBox.show('错误', f'服务器类型 "{data["server_type"]}" 已存在', 'warning', parent=self)
                        return
                    # 删除旧的服务器配置
                    self.db.remove_smtp_server(server_type)
                
                # 添加/更新服务器配置
                if self.db.add_smtp_server(data['server_type'], data['config']):
                    self.load_server_list()
                    self.server_updated.emit()
                    MessageBox.show('成功', '服务器已更新', 'info', parent=self)
                else:
                    raise Exception("更新服务器配置失败")
                
            except Exception as e:
                self.logger.error(f"更新服务器失败: {str(e)}")
                MessageBox.show('错误', f'更新失败：{str(e)}', 'critical', parent=self)
        
    def delete_server(self):
        """删除服务器"""
        current_item = self.server_list.currentItem()
        if not current_item:
            MessageBox.show('提示', '请先选择要删除的服务器', 'warning', parent=self)
            return
        
        server_type = current_item.text()
        
        # 检查是否是默认服务器
        if server_type in self.default_servers:
            MessageBox.show('提示', '默认邮箱服务器不能删除', 'warning', parent=self)
            return
            
        if MessageBox.confirm('确认', f'确定要删除服务器"{server_type}"吗？', parent=self):
            try:
                self.db.remove_smtp_server(server_type)
                self.load_server_list()
                self.server_updated.emit()
                MessageBox.show('成功', '服务器已删除', 'info', parent=self)
            except Exception as e:
                MessageBox.show('错误', f'删除失败：{str(e)}', 'critical', parent=self)
        
    def on_server_selected(self, row):
        """当选择服务器时"""
        if row >= 0:
            server_type = self.server_list.item(row).text()
            server_config = self.db.get_smtp_server(server_type)
            if server_config:
                # 可以在这里显示服务器详细信息
                pass

class ServerInputDialog(QDialog):
    def __init__(self, parent=None, server_type='', config=None, readonly_type=False):
        super().__init__(parent)
        self.setWindowTitle('SMTP服务器设置')
        self.setStyleSheet(MODERN_STYLE)
        self.setMinimumWidth(400)
        
        layout = QVBoxLayout()
        
        # 服务器类型
        type_layout = QHBoxLayout()
        self.type_input = QLineEdit(server_type)
        self.type_input.setPlaceholderText('如：QQ企业邮箱')
        if readonly_type:
            self.type_input.setReadOnly(True)
            self.type_input.setStyleSheet("""
                QLineEdit {
                    background-color: #f5f5f5;
                    color: #666666;
                }
            """)
        type_layout.addWidget(QLabel('服务器类型:'))
        type_layout.addWidget(self.type_input)
        
        # SMTP服务器
        smtp_layout = QHBoxLayout()
        self.smtp_input = QLineEdit()
        self.smtp_input.setPlaceholderText('如：smtp.exmail.qq.com')
        smtp_layout.addWidget(QLabel('SMTP服务器:'))
        smtp_layout.addWidget(self.smtp_input)
        
        # 端口
        port_layout = QHBoxLayout()
        self.port_input = QLineEdit()
        self.port_input.setPlaceholderText('如：465')
        port_layout.addWidget(QLabel('端口:'))
        port_layout.addWidget(self.port_input)
        
        # SSL和TLS选项
        ssl_tls_layout = QHBoxLayout()
        self.ssl_checkbox = QCheckBox('使用SSL')
        self.tls_checkbox = QCheckBox('使用TLS')
        
        # 添加SSL和TLS的互斥逻辑
        self.ssl_checkbox.stateChanged.connect(self.on_ssl_changed)
        self.tls_checkbox.stateChanged.connect(self.on_tls_changed)
        
        ssl_tls_layout.addWidget(self.ssl_checkbox)
        ssl_tls_layout.addWidget(self.tls_checkbox)
        
        # 按钮
        btn_layout = QHBoxLayout()
        save_btn = QPushButton('保存')
        save_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton('取消')
        cancel_btn.clicked.connect(self.reject)
        
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        
        # 如果有现有配置，填充表单
        if config:
            self.smtp_input.setText(config.get('smtp_server', ''))
            self.port_input.setText(str(config.get('smtp_port', '')))
            self.ssl_checkbox.setChecked(config.get('use_ssl', True))
            self.tls_checkbox.setChecked(config.get('use_tls', False))
        
        # 添加所有布局
        layout.addLayout(type_layout)
        layout.addLayout(smtp_layout)
        layout.addLayout(port_layout)
        layout.addLayout(ssl_tls_layout)
        layout.addLayout(btn_layout)
        
        self.setLayout(layout)
    
    def on_ssl_changed(self, state):
        """当SSL选项改变时"""
        if state == Qt.Checked:
            self.tls_checkbox.setChecked(False)
    
    def on_tls_changed(self, state):
        """当TLS选项改变时"""
        if state == Qt.Checked:
            self.ssl_checkbox.setChecked(False)
    
    def get_data(self):
        """获取表单数据"""
        return {
            'server_type': self.type_input.text().strip(),
            'config': {
                'smtp_server': self.smtp_input.text().strip(),
                'smtp_port': int(self.port_input.text().strip() or 0),
                'use_ssl': self.ssl_checkbox.isChecked(),
                'use_tls': self.tls_checkbox.isChecked()
            }
        } 