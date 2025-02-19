import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                            QTextEdit, QFileDialog, QProgressBar, QComboBox, QDialog, QSplitter, QTextBrowser, QTableWidget, QTableWidgetItem, QHeaderView)
from PyQt5.QtCore import Qt, pyqtSignal, QTimer
from PyQt5.QtGui import QColor, QBrush, QIcon
import pandas as pd
from email_utils import EmailServer
from template_dialog import TemplateDialog
from sender_dialog import SenderDialog
from styles import MODERN_STYLE
from database import Database
from server_dialog import ServerDialog
from config import Config
from single_instance import SingleInstance
from variable_dialog import VariableDialog
from datetime import datetime
from log_dialog import LogDialog
from logger import Logger
from system_log_dialog import SystemLogDialog

def get_resource_path(relative_path):
    """获取资源文件的绝对路径
    
    Args:
        relative_path: 相对于程序根目录的路径
        
    Returns:
        str: 资源文件的绝对路径
    """
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller 打包后的路径
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

class EmailSender(QMainWindow):
    activate_window_signal = pyqtSignal()
    
    def __init__(self):
        super().__init__()
        # 连接激活窗口信号
        self.activate_window_signal.connect(self.activate_window)
        
        # 单例模式检查
        self.instance = SingleInstance(self)
        if self.instance.already:
            self.instance.find_and_activate_window()
            sys.exit(0)

        # 基础配置初始化
        self.config = Config()
        self.db = Database(self.config)
        # 确保数据库初始化
        self.db.init_database()
        
        # 初始化日志系统
        self.logger = Logger()
        Logger.init_db(self.db)  # 设置数据库连接
        self.logger.info("邮件发送程序启动")
        
        # 设置窗口基本属性
        icon_path = get_resource_path('imgs/logo.ico')
        self.setWindowIcon(QIcon(icon_path))
        self.setStyleSheet(MODERN_STYLE)
        
        # 初始化基本UI
        self.initUI()
        
        # 延迟加载数据
        QTimer.singleShot(100, self.load_data)

    def load_data(self):
        """延迟加载数据"""
        self.load_templates()
        self.load_last_sender()
        
    def initUI(self):
        """初始化UI"""
        # 设置基本窗口属性
        self.setWindowTitle('卓帆 - 邮件发送工具')
        # 设置窗口大小

        self.setGeometry(100, 100, 1200, 700)
        
        # 创建基本UI结构
        self.create_menu_bar()
        self.create_main_layout()
        
    def create_menu_bar(self):
        """创建菜单栏"""
        menubar = self.menuBar()
        
        # 设置菜单栏样式
        menubar.setStyleSheet("""
            QMenuBar {
                background-color: #f8f9fa;
                font-family: "Microsoft YaHei", "Segoe UI";
                font-size: 14px;  /* 增大字体大小 */
                padding: 2px;
            }
            QMenuBar::item {
                padding: 6px 10px;
                margin: 0px;
                background: transparent;
            }
            QMenuBar::item:selected {
                background-color: #e6f7ff;
                border-radius: 3px;
            }
            QMenu {
                background-color: white;
                font-family: "Microsoft YaHei", "Segoe UI";
                font-size: 14px;  /* 增大下拉菜单字体大小 */
                padding: 4px;
                border: 1px solid #e8e8e8;
                border-radius: 3px;
            }
            QMenu::item {
                padding: 6px 20px;
                border-radius: 3px;
            }
            QMenu::item:selected {
                background-color: #e6f7ff;
                color: #333333;
            }
        """)
        
        # 文件菜单
        file_menu = menubar.addMenu('文件')
        import_excel_action = file_menu.addAction('导入Excel')
        import_excel_action.triggered.connect(self.import_excel)
        file_menu.addSeparator()
        exit_action = file_menu.addAction('退出')
        exit_action.triggered.connect(self.close)
        
        # 工具菜单
        tools_menu = menubar.addMenu('工具')
        manage_sender_action = tools_menu.addAction('发件人管理')
        manage_sender_action.triggered.connect(self.manage_senders)
        mail_server_action = tools_menu.addAction('服务器管理')
        mail_server_action.triggered.connect(self.manage_mail_server)
        manage_template_action = tools_menu.addAction('模板管理')
        manage_template_action.triggered.connect(self.add_template)
        manage_variable_action = tools_menu.addAction('变量管理')
        manage_variable_action.triggered.connect(self.manage_variables)
        tools_menu.addSeparator()  # 添加分隔线
        history_action = tools_menu.addAction('发送历史记录')
        history_action.triggered.connect(self.show_history)
        system_log_action = tools_menu.addAction('系统日志')
        system_log_action.triggered.connect(self.show_system_log)
        
        # 帮助菜单
        help_menu = menubar.addMenu('帮助')
        about_action = help_menu.addAction('关于')
        about_action.triggered.connect(self.show_about)

    def import_excel(self):
        """导入Excel文件"""
        self.logger.info("开始导入Excel文件")
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "选择Excel文件",
            "",
            "Excel Files (*.xlsx *.xls)"
        )
        
        if not file_name:
            self.status_label.setText("用户取消导入Excel")
            self.logger.info("用户取消导入Excel")
            return
        
        try:
            self.df = pd.read_excel(file_name)
            self.logger.info(f"成功导入Excel文件: {file_name}")
            required_columns = ['姓名', '收件人邮箱']
            
            # 检查必需列
            if not all(col in self.df.columns for col in required_columns):
                self.status_label.setText('Excel文件必须包含：姓名、收件人邮箱')
                self.df = None
                return
            
            # 将Excel表头添加到变量列表中
            for column in self.df.columns:
                try:
                    # 检查变量是否已存在
                    variables = self.db.get_variables()
                    existing_vars = [var['name'] for var in variables]
                    
                    if column not in existing_vars:
                        self.db.add_variable(column)
                        self.status_label.setText(f'已添加新变量：{column}')
                except Exception as e:
                    self.status_label.setText(f'添加变量失败：{str(e)}')
            
            # 清空并更新日志表格
            self.log_table.setRowCount(0)
            for index, row in self.df.iterrows():
                row_position = self.log_table.rowCount()
                self.log_table.insertRow(row_position)
                
                # 创建表格项并设置居中对齐
                name_item = QTableWidgetItem(str(row['姓名']))
                name_item.setTextAlignment(Qt.AlignCenter)
                
                email_item = QTableWidgetItem(str(row['收件人邮箱']))
                email_item.setTextAlignment(Qt.AlignCenter)
                
                status_item = QTableWidgetItem('等待发送')
                status_item.setTextAlignment(Qt.AlignCenter)
                status_item.setForeground(QBrush(QColor('#666666')))  # 灰色
                
                # 设置表格项
                self.log_table.setItem(row_position, 0, name_item)
                self.log_table.setItem(row_position, 1, email_item)
                self.log_table.setItem(row_position, 2, status_item)
            
            # 发出变量更新信号，更新所有相关UI
            self.on_variable_updated()
            
            self.status_label.setText(f'已导入 {len(self.df)} 条数据，并更新变量列表')
            
        except pd.errors.ParserError:
            self.status_label.setText('Excel文件格式错误')
            self.df = None
        except Exception as e:
            self.logger.error(f"导入Excel失败: {str(e)}")
            self.status_label.setText(f'导入失败：{str(e)}')
            self.df = None

    def preview_email(self):
        if not hasattr(self, 'df') or self.df is None or len(self.df.index) == 0:
            self.status_label.setText('请先到文件菜单导入Excel文件')
            return
            
        try:
            # 创建预览窗口
            preview_window = QWidget()
            preview_window.setWindowTitle('邮件预览')
            preview_window.setGeometry(200, 200, 600, 400)
            preview_layout = QVBoxLayout(preview_window)
            
            # 检查是否有模板内容
            template_name = self.template_combo.currentText()
            if not template_name:
                self.status_label.setText('请先选择邮件模板')
                preview_window.close()
                return
                
            templates = self.db.get_templates()
            template = templates.get(template_name)
            if not template:
                self.status_label.setText('模板加载失败')
                preview_window.close()
                return
            
            # 预览标题
            try:
                preview_title = template['title'].format(**self.df.iloc[0])
                title_label = QLabel(f"邮件标题: {preview_title}")
                preview_layout.addWidget(title_label)
            except Exception as e:
                self.status_label.setText(f'标题预览失败：{str(e)}')
                preview_window.close()
                return
                
            # 预览正文
            try:
                # 获取HTML内容并进行变量替换
                content_html = template['content']
                # 先处理变量替换
                data_dict = self.df.iloc[0].to_dict()
                for key, value in data_dict.items():
                    placeholder = '{' + key + '}'
                    content_html = content_html.replace(placeholder, str(value))
                
                # 创建预览文本框并设置HTML内容
                preview_text = QTextEdit()
                preview_text.setHtml(content_html)
                preview_text.setReadOnly(True)
                
                preview_layout.addWidget(QLabel("邮件正文:"))
                preview_layout.addWidget(preview_text)
                
                # 显示收件人信息
                info_text = (f"预览收件人: {self.df.iloc[0]['姓名']}\n"
                            f"预览邮箱: {self.df.iloc[0]['收件人邮箱']}")
                preview_layout.addWidget(QLabel(info_text))
                
                # 添加关闭按钮
                close_btn = QPushButton("关闭预览")
                close_btn.clicked.connect(preview_window.close)
                preview_layout.addWidget(close_btn)
                
                preview_window.show()
                
                # 保持窗口引用，防止被垃圾回收
                self.preview_window = preview_window
                
            except Exception as e:
                self.status_label.setText(f'预览失败：{str(e)}')
                preview_window.close()
                
        except Exception as e:
            self.status_label.setText(f'预览窗口创建失败：{str(e)}')

    def send_test_email(self):
        """发送测试邮件到发件人邮箱"""
        self.logger.info("开始发送测试邮件")
        # 检查是否选择了模板
        template_name = self.template_combo.currentText()
        if not template_name:
            self.status_label.setText('请先选择邮件模板')
            return

        if not self.current_sender:
            self.status_label.setText('请选择发件人')
            return

        try:
            # 从数据库获取模板
            templates = self.db.get_templates()
            template = templates.get(template_name)
            if not template:
                self.status_label.setText('模板加载失败')
                return

            # 从配置文件获取测试数据
            test_data = self.config.get_test_data()
            # 添加收件人邮箱
            test_data['收件人邮箱'] = self.current_sender['email']

            # 使用配置的邮箱类型
            email_server = EmailServer(
                self.current_sender['email'],
                self.current_sender['password'],
                self.current_sender.get('server_type', 'QQ企业邮箱')
            )
            email_server.connect()

            # 替换模板中的变量
            content = template['content']
            for key, value in test_data.items():
                placeholder = '{' + key + '}'
                content = content.replace(placeholder, str(value))
            subject = template['title'].format(**test_data)
            
            success, message = email_server.send_email(
                test_data['收件人邮箱'],
                subject,
                content
            )

            if success:
                self.status_label.setText('测试邮件发送成功')
                self.logger.info("测试邮件发送成功")
            else:
                self.status_label.setText(f'测试邮件发送失败: {message}')

            email_server.close()

        except Exception as e:
            self.logger.error(f"测试邮件发送失败: {str(e)}")
            self.status_label.setText(f'发送失败：{str(e)}')

    def start_sending(self):
        """开始批量发送邮件"""
        self.logger.info("开始批量发送邮件")
        if not self.check_send_conditions():
            return
        
        try:
            # 准备发送
            template_name = self.template_combo.currentText()
            templates = self.db.get_templates()  # 从数据库获取模板
            template = templates.get(template_name)
            if not template:
                self.status_label.setText('模板不存在')
                return
            
            # 设置进度条
            total = len(self.df)
            self.progress_bar.setMaximum(total)
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat('正在发送 - %p% (%v/%m)')
            
            # 创建邮件服务器连接
            server = EmailServer(
                self.current_sender['email'],
                self.current_sender['password'],
                self.current_sender.get('server_type')
            )
            
            # 遍历发送
            for index, row in self.df.iterrows():
                try:
                    # 获取测试数据
                    test_data = self.config.get_test_data()
                    # 更新数据，Excel数据优先
                    test_data.update(row.to_dict())
                    
                    # 替换变量
                    title = template['title']
                    content = template['content']
                    for key, value in test_data.items():
                        placeholder = '{' + key + '}'
                        title = title.replace(placeholder, str(value))
                        content = content.replace(placeholder, str(value))
                    
                    # 发送邮件
                    success, message = server.send_email(
                        test_data['收件人邮箱'],
                        title,
                        content
                    )
                    
                    # 更新状态
                    status_item = self.log_table.item(index, 2)
                    if success:
                        status_item.setText('发送成功')
                        status_item.setForeground(QBrush(QColor('#28a745')))  # 绿色
                    else:
                        status_item.setText(f'发送失败: {message}')
                        status_item.setForeground(QBrush(QColor('#dc3545')))  # 红色
                    
                    # 更新进度条
                    self.progress_bar.setValue(index + 1)
                    if index + 1 == total:
                        self.progress_bar.setFormat('发送完成 - 100% (%m/%m)')
                    QApplication.processEvents()
                    
                    # 记录发送日志
                    self.db.add_send_log(
                        batch_id=f"BATCH_{datetime.now().strftime('%Y%m%d%H%M%S')}",
                        sender_email=self.current_sender['email'],
                        recipient_email=test_data['收件人邮箱'],
                        recipient_name=test_data['姓名'],
                        subject=title,
                        status='成功' if success else '失败',
                        error_message=message if not success else None
                    )
                    
                except Exception as e:
                    # 更新失败状态
                    status_item = self.log_table.item(index, 2)
                    status_item.setText(f'发送失败: {str(e)}')
                    status_item.setForeground(QBrush(QColor('#dc3545')))  # 红色
                
            server.close()
            self.status_label.setText('发送完成')
            self.logger.info(f"批量发送完成，共发送{total}封邮件")
            
            # 如果是最后一封邮件,自动导出结果
            if index + 1 == total:
                self.export_batch_result(f"BATCH_{datetime.now().strftime('%Y%m%d%H%M%S')}")
            
        except Exception as e:
            self.progress_bar.setFormat('发送失败 - %p% (%v/%m)')
            self.logger.error(f"批量发送失败: {str(e)}")
            self.status_label.setText(f'发送失败：{str(e)}')

    def load_templates(self):
        """加载模板列表"""
        self.template_combo.clear()
        templates = self.db.get_templates()  # 从数据库获取模板
        for name in templates.keys():
            self.template_combo.addItem(name)
        
        # 更新预览
        self.update_preview()

    def update_preview(self):
        """更新预览内容"""
        template_name = self.template_combo.currentText()
        if not template_name:
            self.preview_title.clear()
            self.preview_content.clear()
            return
        
        try:
            templates = self.db.get_templates()  # 从数据库获取模板
            template = templates.get(template_name)
            if template:
                # 如果有Excel数据，使用第一行数据，否则使用测试数据
                if hasattr(self, 'df') and self.df is not None and len(self.df) > 0:
                    test_data = self.df.iloc[0].to_dict()
                else:
                    test_data = self.config.get_test_data()
                
                # 替换标题中的变量
                title = template['title']
                for var, value in test_data.items():
                    title = title.replace(f'{{{var}}}', str(value))
                    
                # 替换内容中的变量
                content = template['content']
                for var, value in test_data.items():
                    content = content.replace(f'{{{var}}}', str(value))
                    
                # 更新预览
                self.preview_title.setText(title)
                self.preview_content.setHtml(content)
        except Exception as e:
            self.status_label.setText(f'加载模板失败: {str(e)}')

    def add_template(self):
        """打开添加模板对话框"""
        dialog = TemplateDialog(self)
        
        # 计算对话框相对于主窗口的居中位置
        x = self.x() + (self.width() - dialog.width()) // 2
        y = self.y() + (self.height() - dialog.height()) // 2
        dialog.move(x, y)
        
        # 连接模板更新信号
        dialog.template_updated.connect(self.load_templates)
        dialog.exec_()

    def preview_template(self, template_name):
        """当选择模板时更新预览"""
        if not template_name:
            self.preview_content.clear()
            self.status_label.setText('未选择模板')
            return
        
        try:
            # 从数据库加载模板
            templates = self.db.get_templates()
            template = templates.get(template_name)
            
            if template:
                # 如果有Excel数据，使用第一行数据，否则使用测试数据
                if hasattr(self, 'df') and self.df is not None and len(self.df) > 0:
                    test_data = self.df.iloc[0].to_dict()
                else:
                    test_data = self.config.get_test_data()
                    if not test_data:  # 如果没有测试数据，创建默认值
                        variables = self.db.get_variables()
                        test_data = {var: f'测试{var}' for var in variables}
                        self.config.save_test_data(test_data)
                
                # 替换标题中的变量
                title = template['title']
                for key, value in test_data.items():
                    placeholder = '{' + key + '}'
                    title = title.replace(placeholder, str(value))
                
                # 替换内容中的变量
                content = template['content']
                for key, value in test_data.items():
                    placeholder = '{' + key + '}'
                    content = content.replace(placeholder, str(value))
                
                # 更新预览
                self.preview_title.setText(title)
                self.preview_content.setHtml(content)
                self.status_label.setText(f'已加载"{template_name}"模板')
        except Exception as e:
            self.status_label.setText(f'加载模板失败: {str(e)}')

    def on_sender_changed(self, index):
        """当选择发件人时更新当前发件人信息"""
        if index < 0:
            self.current_sender = None
            return
            
        # 直接从combobox的item data中获取sender数据
        self.current_sender = self.sender_combo.itemData(index)
        
        # 保存最后使用的发件人
        if self.current_sender:
            self.config.save_last_sender(self.current_sender['email'])

    def load_last_sender(self):
        """加载发件人列表并选择最后一个"""
        self.update_sender_list()
        
        # 尝试恢复上次使用的发件人
        last_sender = self.config.get_last_sender()
        if last_sender:
            index = self.sender_combo.findText(last_sender)
            if index >= 0:
                self.sender_combo.setCurrentIndex(index)

    def manage_senders(self):
        """打开发件人管理对话框"""
        dialog = SenderDialog(self)
        
        # 计算对话框相对于主窗口的居中位置
        x = self.x() + (self.width() - dialog.width()) // 2
        y = self.y() + (self.height() - dialog.height()) // 2
        dialog.move(x, y)
        
        # 设置更新回调
        dialog.sender_updated.connect(self.update_sender_list)
        if dialog.exec_() == QDialog.Accepted:
            self.update_sender_list()

    def update_sender_list(self):
        """更新发件人下拉列表"""
        current_email = self.sender_combo.currentText()  # 保存当前选中的邮箱
        self.sender_combo.clear()
        
        # 加载发件人列表
        senders = self.config.get_sender_list()
        for sender in senders:
            self.sender_combo.addItem(sender['email'], sender)  # 将整个sender数据作为item的data
            
        # 尝试恢复之前选中的邮箱
        if current_email:
            index = self.sender_combo.findText(current_email)
            if index >= 0:
                self.sender_combo.setCurrentIndex(index)
            elif self.sender_combo.count() > 0:  # 如果找不到之前的邮箱，选择第一个
                self.sender_combo.setCurrentIndex(0)
        elif self.sender_combo.count() > 0:  # 如果没有之前选中的邮箱，选择第一个
            self.sender_combo.setCurrentIndex(0)
        else:
            self.current_sender = None

    def handle_link_click(self, url):
        """处理链接点击"""
        import webbrowser
        webbrowser.open(url.toString())
        # 重新加载预览内容，防止被清空
        template_name = self.template_combo.currentText()
        if template_name:
            try:
                # 从数据库加载模板
                templates = self.db.get_templates()
                template = templates.get(template_name)
                if template:
                    # 使用config/config.ini中的测试数据
                    test_data = self.config.get_test_data()
                    
                    # 替换标题中的变量
                    title = template['title']
                    for var, value in test_data.items():
                        title = title.replace(f'{{{var}}}', value)
                        
                    # 替换内容中的变量
                    content = template['content']
                    for var, value in test_data.items():
                        content = content.replace(f'{{{var}}}', value)
                        
                    # 更新预览
                    self.preview_title.setText(title)
                    self.preview_content.setHtml(content)
            except Exception as e:
                self.status_label.setText(f'加载模板失败: {str(e)}')

    def manage_mail_server(self):
        """打开邮件服务器管理对话框"""
        dialog = ServerDialog(self)
        
        # 计算对话框相对于主窗口的居中位置
        x = self.x() + (self.width() - dialog.width()) // 2
        y = self.y() + (self.height() - dialog.height()) // 2
        dialog.move(x, y)
        
        dialog.exec_()

    def manage_variables(self):
        """打开变量管理对话框"""
        dialog = VariableDialog(self)
        dialog.variable_updated.connect(self.on_variable_updated)
        dialog.exec_()

    def on_variable_updated(self):
        """当变量更新时的处理方法"""
        # 重新加载模板，因为模板中的变量可能已更新
        self.load_templates()

    def check_send_conditions(self):
        """检查发送条件"""
        if not hasattr(self, 'df') or self.df is None or len(self.df.index) == 0:
            self.status_label.setText('请在文件菜单 - 导入Excel文件')
            return False
        
        if not self.current_sender:
            self.status_label.setText('请在工具菜单 - 发件人管理中添加发件人')
            return False

        if not self.template_combo.currentText():
            self.status_label.setText('请在工具菜单 - 模板管理中添加邮件模板')
            return False

        return True

    def show_about(self):
        """显示关于对话框"""
        about_dialog = QDialog(self)
        about_dialog.setWindowTitle('关于')
        # 设置对话框大小根据文本自适应
        about_dialog.resize(400, 600) # 设置对话框的固定大小
        
        # 计算对话框相对于主窗口的居中位置
        x = self.x() + (self.width() - about_dialog.width()) // 2
        y = self.y() + (self.height() - about_dialog.height()) // 2
        about_dialog.move(x, y)
        
        layout = QVBoxLayout(about_dialog)
        layout.setContentsMargins(20, 20, 20, 20) # 设置对话框的边距
        layout.setSpacing(15)
        
        # 修改图标加载方式
        icon_label = QLabel()
        icon_path = get_resource_path('imgs/logo.ico')
        icon_label.setPixmap(QIcon(icon_path).pixmap(64, 64))
        icon_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(icon_label)
        
        # 添加文本
        text_label = QLabel("""
            <div style='text-align: center;'>
                <h2 style='color: #40a9ff;'>邮件批量发送程序</h2>
                <p style='margin: 15px 0; line-height: 1.5;'>
                    这是一个专业的批量邮件发送工具，支持多种邮箱服务商，
                    可以自定义邮件模板，支持变量替换，让邮件发送更加高效便捷。
                </p>
                <p style='margin: 10px 0;'>
                    <b>主要功能：</b>
                </p>
                <p style='margin: 5px 0; line-height: 1.5;'>
                    • 支持多种邮箱服务商（QQ邮箱、企业邮箱等）<br>
                    • 自定义邮件模板，支持HTML格式<br>
                    • Excel导入收件人信息<br>
                    • 实时发送状态监控<br>
                    • 批量发送进度显示
                </p>
                <p style='margin: 10px 0;'><b>版本：</b>1.0.0</p>
                <p style='margin: 10px 0;'><b>开发商：</b>BBRHub</p>
                <p style='margin: 10px 0;'><b>版权所有：</b>© 2025 保留所有权利</p>
            </div>
        """)
        text_label.setAlignment(Qt.AlignCenter)
        text_label.setWordWrap(True)  # 允许文本换行
        layout.addWidget(text_label)
        
        # 添加关闭按钮
        close_btn = QPushButton('关闭')
        close_btn.setFixedWidth(100) # 设置按钮宽度
        close_btn.clicked.connect(about_dialog.close)
        
        btn_layout = QHBoxLayout()
        btn_layout.addWidget(close_btn)
        btn_layout.setAlignment(Qt.AlignCenter)
        layout.addLayout(btn_layout)
        
        # 禁用对话框的系统菜单，防止拖动
        about_dialog.setWindowFlags(Qt.Dialog | Qt.MSWindowsFixedSizeDialogHint)
        
        about_dialog.exec_()

    def show_log_tooltip(self, row, column):
        """显示日志表格的工具提示"""
        item = self.log_table.item(row, column)
        if not item:
            return
        
        # 获取该行的所有信息
        name = self.log_table.item(row, 0).text() if self.log_table.item(row, 0) else ''
        email = self.log_table.item(row, 1).text() if self.log_table.item(row, 1) else ''
        status = self.log_table.item(row, 2).text() if self.log_table.item(row, 2) else ''
        
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
                <p style='margin: 0;'><b>姓名:</b> {name}</p>
                <p style='margin: 5px 0;'><b>邮箱:</b> {email}</p>
                <p style='margin: 0;'><b>状态:</b> 
                    <span style='color: {
                        "#52c41a" if "成功" in status 
                        else "#ff4d4f" if "失败" in status 
                        else "#faad14"
                    };'>{status}</span>
                </p>
            </div>
        """
        
        # 设置工具提示
        item.setToolTip(tooltip)

    def add_log_entry(self, name, email, status):
        """添加日志条目"""
        row = self.log_table.rowCount()
        self.log_table.insertRow(row)
        
        # 创建并设置单元格项，使其居中对齐
        items = [
            QTableWidgetItem(str(name)),
            QTableWidgetItem(str(email)),
            QTableWidgetItem(str(status))
        ]
        
        for col, item in enumerate(items):
            item.setTextAlignment(Qt.AlignCenter)
            self.log_table.setItem(row, col, item)

    def activate_window(self):
        """激活窗口"""
        # 确保窗口不是最小化的
        self.showNormal()
        # 将窗口置于最前
        self.setWindowState(self.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
        # 激活窗口
        self.activateWindow()
        # 将窗口置顶
        self.raise_()

    def create_main_layout(self):
        """创建主界面布局"""
        # 创建中心部件和主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QHBoxLayout(central_widget)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(20)

        # 左侧控制和日志区域
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setSpacing(15)
        left_layout.setContentsMargins(0, 0, 0, 0)

        # 发件人组
        sender_group = QWidget()
        sender_group.setObjectName("groupBox")
        sender_layout = QHBoxLayout(sender_group)
        sender_layout.setContentsMargins(10, 10, 10, 10)
        
        self.sender_combo = QComboBox()
        self.sender_combo.setMinimumWidth(200)
        self.sender_combo.currentIndexChanged.connect(self.on_sender_changed)
        
        sender_layout.addWidget(QLabel('发件人:'))
        sender_layout.addWidget(self.sender_combo, 1)

        # 模板选择组
        template_group = QWidget()
        template_group.setObjectName("groupBox")
        template_layout = QHBoxLayout(template_group)
        template_layout.setContentsMargins(10, 10, 10, 10)
        
        self.template_combo = QComboBox()
        self.template_combo.setMinimumWidth(200)
        self.template_combo.currentTextChanged.connect(self.preview_template)
        
        template_layout.addWidget(QLabel('模板:'))
        template_layout.addWidget(self.template_combo, 1)

        # 发送按钮组
        button_group = QWidget()
        button_group.setObjectName("groupBox")
        button_layout = QHBoxLayout(button_group)
        button_layout.setContentsMargins(10, 10, 10, 10)
        
        self.test_btn = QPushButton('发送测试')
        self.test_btn.clicked.connect(self.send_test_email)
        self.send_btn = QPushButton('开始发送')
        self.send_btn.clicked.connect(self.start_sending)
        
        button_layout.addWidget(self.test_btn)
        button_layout.addWidget(self.send_btn)

        # 进度条组
        progress_group = QWidget()
        progress_group.setObjectName("groupBox")
        progress_layout = QVBoxLayout(progress_group)
        progress_layout.setContentsMargins(10, 10, 10, 10)
        
        progress_label = QLabel('发送进度:')
        self.progress_bar = QProgressBar()
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.progress_bar.setMinimumHeight(25)
        self.progress_bar.setFormat('等待发送 - %p% (%v/%m)')
        
        progress_layout.addWidget(progress_label)
        progress_layout.addWidget(self.progress_bar)

        # 状态标签
        self.status_label = QLabel('就绪')
        self.status_label.setObjectName('status_label')
        self.status_label.setMinimumHeight(30)

        # 发送日志区域
        log_group = QWidget()
        log_group.setObjectName("groupBox")
        log_layout = QVBoxLayout(log_group)
        log_layout.setContentsMargins(10, 10, 10, 10)
        
        log_layout.addWidget(QLabel('发送日志:'))
        self.log_table = QTableWidget()
        self.log_table.setColumnCount(3)
        self.log_table.setHorizontalHeaderLabels(['姓名', '邮箱', '发送状态'])
        
        # 修改表格列的拉伸模式，确保内容居中
        header = self.log_table.horizontalHeader()
        for i in range(3):
            header.setSectionResizeMode(i, QHeaderView.Stretch)
        
        # 设置表格样式，强化居中显示
        self.log_table.setStyleSheet("""
            QTableWidget {
                background-color: white;
                alternate-background-color: #fafafa;
            }
            QTableWidget::item {
                padding: 5px;
                text-align: center;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                padding: 5px;
                text-align: center;
            }
        """)
        
        # 设置默认对齐方式为居中
        self.log_table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        
        # 其他设置保持不变
        self.log_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.log_table.setSelectionMode(QTableWidget.NoSelection)
        self.log_table.verticalHeader().setVisible(False)
        self.log_table.setAlternatingRowColors(True)
        self.log_table.setMouseTracking(True)
        self.log_table.setToolTipDuration(5000)
        
        # 连接单元格进入事件
        self.log_table.cellEntered.connect(self.show_log_tooltip)
        
        log_layout.addWidget(self.log_table)

        # 添加组件到左侧布局
        left_layout.addWidget(sender_group)
        left_layout.addWidget(template_group)
        left_layout.addWidget(progress_group)
        left_layout.addWidget(button_group)
        left_layout.addWidget(self.status_label)
        left_layout.addWidget(log_group, 1)

        # 右侧预览区域
        preview_group = QWidget()
        preview_group.setObjectName("groupBox")
        preview_layout = QVBoxLayout(preview_group)
        preview_layout.setContentsMargins(10, 10, 10, 10)
        preview_layout.setSpacing(10)

        # 预览标题
        title_layout = QHBoxLayout()
        title_label = QLabel('邮件标题:')
        title_label.setFixedWidth(80)
        self.preview_title = QLineEdit()
        self.preview_title.setReadOnly(True)
        title_layout.addWidget(title_label)
        title_layout.addWidget(self.preview_title)

        # 预览内容
        content_label = QLabel('邮件内容:')
        self.preview_content = QTextBrowser()
        self.preview_content.setOpenExternalLinks(False)
        self.preview_content.setOpenLinks(False)
        self.preview_content.setMinimumWidth(600)
        self.preview_content.setStyleSheet("""
            QTextBrowser {
                padding: 10px;
                line-height: 1.5;
            }
        """)

        preview_layout.addLayout(title_layout)
        preview_layout.addWidget(content_label)
        preview_layout.addWidget(self.preview_content)

        # 使用分割器添加左右两侧
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(left_widget)
        splitter.addWidget(preview_group)
        splitter.setStretchFactor(0, 2)  # 左侧比例为2
        splitter.setStretchFactor(1, 4)  # 右侧比例为4

        # 设置初始大小
        width = self.width()
        splitter.setSizes([int(width * 0.333), int(width * 0.667)])  # 保持1.5:3的比例

        # 禁止调整分割器
        splitter.handle(1).setEnabled(False)  # 禁用分割器手柄
        splitter.setHandleWidth(0)  # 设置分割线宽度为0，使其不可见

        # 设置分割器的样式，隐藏分割线
        splitter.setStyleSheet("""
            QSplitter::handle {
                background: none;
                border: none;
            }
        """)

        # 主布局
        main_layout = QVBoxLayout()
        main_layout.addWidget(splitter)

        # 添加版权信息
        copyright_bar = QWidget()
        copyright_bar.setFixedHeight(30)
        copyright_bar.setStyleSheet("""
            QWidget {
                background-color: transparent;
            }
            QLabel {
                color: #666666;
                font-size: 12px;
            }
        """)
        copyright_layout = QHBoxLayout(copyright_bar)
        copyright_layout.setContentsMargins(10, 0, 10, 0)

        copyright_label = QLabel('© 2025 BBRHub All Rights Reserved.')
        copyright_label.setAlignment(Qt.AlignCenter)
        copyright_layout.addWidget(copyright_label)
        
        # 将版权信息添加到主布局
        main_layout.addWidget(copyright_bar)
        # 设置中心部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        central_widget.setLayout(main_layout)

    def export_batch_result(self, batch_id):
        """导出批次发送结果"""
        try:
            logs = self.db.get_send_logs(batch_id=batch_id)
            # 创建一个列名映射，确保列数匹配
            df = pd.DataFrame(logs, columns=[
                'id',
                'batch_id',
                'sender_email',
                'recipient_email',
                'recipient_name',
                'subject',
                'status',
                'error_message',
                'send_time'
            ])
            
            # 重新组织要导出的列
            export_df = df[[
                'send_time',
                'sender_email',
                'recipient_email',
                'recipient_name',
                'subject',
                'status',
                'error_message'
            ]]
            
            # 重命名列
            export_df.columns = [
                '发送时间',
                '发件人',
                '收件人',
                '收件人姓名',
                '邮件主题',
                '发送状态',
                '错误信息'
            ]
            
            file_name = f'发送结果_{batch_id}.xlsx'
            export_df.to_excel(file_name, index=False)
            
            self.status_label.setText(f'发送完成,结果已导出到: {file_name}')
        except Exception as e:
            self.status_label.setText(f'导出结果失败: {str(e)}')
    def show_history(self):
        """显示历史记录对话框"""
        dialog = LogDialog(self)
        dialog.exec_()

    def show_system_log(self):
        """显示系统日志对话框"""
        dialog = SystemLogDialog(self)
        dialog.exec_()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    # 修改图标加载方式
    app.setWindowIcon(QIcon(get_resource_path('imgs/logo.ico')))
    ex = EmailSender()
    ex.show()
    sys.exit(app.exec_()) 