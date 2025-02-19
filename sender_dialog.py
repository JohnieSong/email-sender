from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                            QLineEdit, QPushButton, QMessageBox, QListWidget, 
                            QComboBox, QCompleter)
from PyQt5.QtCore import Qt, pyqtSignal, QStringListModel, QTimer
from config import Config
from message_box import MessageBox
from email_utils import EmailServer
from styles import MODERN_STYLE
from database import Database

class SenderInputDialog(QDialog):
    """发件人信息输入对话框"""
    def __init__(self, parent=None, email='', password='', server_type='QQ企业邮箱'):
        super().__init__(parent)
        self.db = Database()  # 获取数据库实例
        self.setWindowTitle('添加发件人')
        self.setStyleSheet(MODERN_STYLE)
        self.setMinimumWidth(400)
        
        # 创建主布局
        layout = QVBoxLayout()
        layout.setSpacing(10)
        
        # 邮箱输入
        email_layout = QHBoxLayout()
        email_label = QLabel('邮箱:')
        email_label.setFixedWidth(60)
        self.email_input = QLineEdit(email)
        self.email_input.setPlaceholderText('请输入邮箱地址')
        if email:  # 如果是编辑模式
            self.email_input.setReadOnly(True)
        
        # 创建邮箱补全列表
        self.email_completer = QCompleter()
        self.email_completer.setCaseSensitivity(Qt.CaseInsensitive)
        self.email_input.setCompleter(self.email_completer)
        
        # 获取常用邮箱域名
        self.email_domains = [
            '@qq.com',
            '@exmail.qq.com',
            '@163.com',
            '@gmail.com',
            '@sina.com',
            '@outlook.com',
            '@mxhichina.com',  # 阿里企业邮箱
            '@126.com',
            '@foxmail.com',
            '@yahoo.com'
        ]
        
        # 添加邮箱输入变化事件
        self.email_input.textChanged.connect(self.on_email_changed)
        
        email_layout.addWidget(email_label)
        email_layout.addWidget(self.email_input)
        
        # 邮箱类型选择
        server_layout = QHBoxLayout()
        server_label = QLabel('邮箱类型:')
        server_label.setFixedWidth(60)
        self.server_combo = QComboBox()
        # 从数据库获取服务器类型列表
        server_types = self.db.get_smtp_server_list()
        self.server_combo.addItems(server_types)
        self.server_combo.setCurrentText(server_type)
        server_layout.addWidget(server_label)
        server_layout.addWidget(self.server_combo)
        
        # 密码输入
        password_layout = QHBoxLayout()
        password_label = QLabel('授权码:')
        password_label.setFixedWidth(60)
        self.password_input = QLineEdit(password)
        self.password_input.setPlaceholderText('输入授权码')
        self.password_input.setEchoMode(QLineEdit.Password)
        password_layout.addWidget(password_label)
        password_layout.addWidget(self.password_input)
        
        # 添加说明文本
        tip_label = QLabel('提示: 授权码不是邮箱密码，需要在邮箱服务商网站单独申请')
        tip_label.setStyleSheet('color: #666666; font-size: 12px;')
        
        # 按钮
        btn_layout = QHBoxLayout()
        save_btn = QPushButton('保存')
        save_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton('取消')
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addStretch()
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        
        # 添加所有布局到主布局
        layout.addLayout(email_layout)
        layout.addLayout(server_layout)
        layout.addLayout(password_layout)
        layout.addWidget(tip_label)
        layout.addStretch()
        layout.addLayout(btn_layout)
        
        self.setLayout(layout)

    def on_email_changed(self, text):
        """当邮箱输入变化时自动选择对应的服务器类型"""
        if '@' in text:
            domain = text.split('@')[1].lower()
            server_type = None
            
            # 域名与服务器类型的映射
            domain_map = {
                'qq.com': 'QQ邮箱',
                'foxmail.com': 'QQ邮箱',
                'exmail.qq.com': 'QQ企业邮箱',
                '163.com': '163邮箱',
                '126.com': '126邮箱',
                'yeah.net': '163邮箱',
                '188.com': '163邮箱',
                'gmail.com': 'Gmail',
                'aliyun.com': '阿里邮箱',
                'mxhichina.com': '阿里企业邮箱',
                '189.cn': '电信邮箱',
                'sohu.com': '搜狐邮箱',
                'sina.com': '新浪邮箱',
                'sina.cn': '新浪邮箱',
                '139.com': '移动邮箱',
                'icloud.com': '苹果邮箱',
                'me.com': '苹果邮箱',
                'outlook.com': 'Outlook',
                'hotmail.com': 'Outlook',
                'live.com': 'Outlook',
                'aol.com': 'AOL邮箱',
                'yandex.com': 'Yandex邮箱'
            }
            
            # 根据域名选择服务器类型
            if domain in domain_map:
                server_type = domain_map[domain]
                # 检查该服务器类型是否在下拉列表中
                index = self.server_combo.findText(server_type)
                if index >= 0:
                    self.server_combo.setCurrentIndex(index)
        
        # 同时更新邮箱建议列表
        self.update_email_suggestions(text)

    def show_tip(self, text):
        """显示提示信息"""
        tip_label = self.findChild(QLabel, 'auto_select_tip')
        if not tip_label:
            tip_label = QLabel(text)
            tip_label.setObjectName('auto_select_tip')
            tip_label.setStyleSheet('color: #28a745; font-size: 12px;')  # 绿色提示文本
            # 找到布局中的提示标签位置
            layout = self.layout()
            # 在服务器类型选择后插入提示
            for i in range(layout.count()):
                if isinstance(layout.itemAt(i).layout(), QHBoxLayout):
                    if layout.itemAt(i).layout().itemAt(0).widget().text() == '邮箱类型:':
                        layout.insertWidget(i + 1, tip_label)
                        break
        else:
            tip_label.setText(text)
            tip_label.setVisible(True)
        
        # 3秒后隐藏提示
        QTimer.singleShot(3000, lambda: tip_label.setVisible(False))

    def update_email_suggestions(self, text):
        """更新邮箱建议列表"""
        if '@' in text:
            # 如果已经输入@，只显示匹配的域名
            prefix = text.split('@')[0]
            suggestions = [prefix + domain for domain in self.email_domains 
                         if domain.startswith('@' + text.split('@')[1])]
        else:
            # 如果还没有输入@，显示所有可能的完整邮箱
            suggestions = [text + domain for domain in self.email_domains]
        
        # 更新补全器的模型
        model = QStringListModel()
        model.setStringList(suggestions)
        self.email_completer.setModel(model)

    def get_data(self):
        """获取表单数据"""
        return {
            'email': self.email_input.text().strip(),
            'password': self.password_input.text().strip(),
            'server_type': self.server_combo.currentText()
        }

class SenderDialog(QDialog):
    sender_updated = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.config = Config()
        self.setWindowTitle('发件人管理')
        self.setStyleSheet(MODERN_STYLE)
        self.setMinimumSize(300, 400)
        self.initUI()
        self.load_senders()

    def initUI(self):
        layout = QVBoxLayout()
        layout.setSpacing(10)
        
        # 发件人列表
        self.sender_list = QListWidget()
        
        # 启用工具提示
        self.sender_list.setMouseTracking(True)
        self.sender_list.setToolTipDuration(5000)
        
        # 按钮组
        btn_layout = QHBoxLayout()
        
        # 创建按钮
        add_btn = QPushButton('添加')
        edit_btn = QPushButton('修改')
        delete_btn = QPushButton('删除')
        
        # 连接按钮事件
        add_btn.clicked.connect(self.add_sender)
        edit_btn.clicked.connect(self.edit_sender)
        delete_btn.clicked.connect(self.delete_sender)
        
        # 将按钮添加到布局中，并设置居中对齐
        btn_layout.addStretch()  # 左侧弹性空间
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(edit_btn)
        btn_layout.addWidget(delete_btn)
        btn_layout.addStretch()  # 右侧弹性空间
        
        # 添加组件到主布局
        layout.addWidget(QLabel('发件人列表:'))
        layout.addWidget(self.sender_list)
        layout.addLayout(btn_layout)
        
        self.setLayout(layout)

    def load_senders(self):
        """加载发件人列表"""
        self.sender_list.clear()
        for sender in self.config.get_sender_list():
            self.sender_list.addItem(sender['email'])

    def add_sender(self):
        """添加新发件人"""
        dialog = SenderInputDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            data = dialog.get_data()
            if not all([data['email'], data['password']]):
                MessageBox.show('错误', '请填写所有字段', 'warning', parent=self)
                return
                
            try:
                # 测试连接
                server = EmailServer(data['email'], data['password'], data['server_type'])
                server.connect()
                server.close()
                
                # 保存发件人信息
                self.config.add_sender(data['email'], data['password'], data['server_type'])
                self.load_senders()
                self.sender_updated.emit()
                MessageBox.show('成功', '发件人邮箱已添加', 'info', parent=self)
            except Exception as e:
                MessageBox.show('错误', str(e), 'critical', parent=self)

    def edit_sender(self):
        """修改发件人"""
        current_item = self.sender_list.currentItem()
        if not current_item:
            MessageBox.show('提示', '请先选择要修改的发件人', 'warning', parent=self)
            return
            
        email = current_item.text()
        sender_info = None
        for sender in self.config.get_sender_list():
            if sender['email'] == email:
                sender_info = sender
                break
                
        if not sender_info:
            return
            
        dialog = SenderInputDialog(
            self,
            email=sender_info['email'],
            password=sender_info['password'],
            server_type=sender_info['server_type']
        )
        
        if dialog.exec_() == QDialog.Accepted:
            data = dialog.get_data()
            if not all([data['email'], data['password']]):
                MessageBox.show('错误', '请填写所有字段', 'warning', parent=self)
                return
                
            try:
                # 测试连接
                server = EmailServer(data['email'], data['password'], data['server_type'])
                server.connect()
                server.close()
                
                # 更新发件人信息
                self.config.remove_sender(email)
                self.config.add_sender(data['email'], data['password'], data['server_type'])
                self.load_senders()
                self.sender_updated.emit()
                MessageBox.show('成功', '发件人信息已更新', 'info', parent=self)
            except Exception as e:
                MessageBox.show('错误', str(e), 'critical', parent=self)

    def delete_sender(self):
        """删除发件人"""
        current_item = self.sender_list.currentItem()
        if not current_item:
            MessageBox.show('提示', '请先选择要删除的发件人', 'warning', parent=self)
            return
            
        email = current_item.text()
        reply = MessageBox.show(
            '确认删除', 
            f'确定要删除发件人"{email}"吗？',
            'question',
            [('是', QMessageBox.AcceptRole), ('否', QMessageBox.RejectRole)],
            parent=self
        )
        
        if reply == 0:
            if self.config.remove_sender(email):
                self.load_senders()
                self.sender_updated.emit()
                MessageBox.show('成功', '发件人已删除', 'info', parent=self)
            else:
                MessageBox.show('错误', '删除失败', 'critical', parent=self)

    def show_sender_tooltip(self, item):
        """显示发件人详细信息的工具提示"""
        if not item:
            return
        
        email = item.text()
        sender_info = None
        
        # 查找对应的发件人信息
        for sender in self.config.get_sender_list():
            if sender['email'] == email:
                sender_info = sender
                break
            
        if sender_info:
            # 构建带HTML样式的工具提示内容
            tooltip = f"""
                <div style='
                    background-color: #ffffff;
                    padding: 10px;
                    border-radius: 4px;
                    border: 1px solid #e8e8e8;
                    font-family: "Microsoft YaHei", sans-serif;
                    font-size: 12px;
                    line-height: 1.5;
                '>
                    <p style='margin: 0;'><b>邮箱地址:</b> {sender_info['email']}</p>
                    <p style='margin: 5px 0;'><b>邮箱类型:</b> {sender_info.get('server_type')}</p>
                    <p style='margin: 0;'><b>状态:</b> <span style='color: #52c41a;'>已配置</span></p>
                </div>
            """
            
            # 设置工具提示
            item.setToolTip(tooltip) 
