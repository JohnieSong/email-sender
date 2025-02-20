import sys
import os
import time
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                            QTextEdit, QFileDialog, QProgressBar, QComboBox, QDialog, QSplitter, QTextBrowser, QTableWidget, QTableWidgetItem, QHeaderView, QListWidget, QMessageBox, QListWidgetItem, QInputDialog)
from PyQt5.QtCore import Qt, pyqtSignal, QTimer, QThread
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
from message_box import MessageBox
from email import message_from_file
from email.header import Header
from bs4 import BeautifulSoup
from about_dialog import AboutDialog
from queue import Queue
import threading
import chardet

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

        # 添加计时器
        self.countdown_timer = QTimer()
        self.countdown_timer.timeout.connect(self.update_test_button)
        self.countdown_timer.setInterval(1000)  # 每秒更新一次

        self.send_thread = None
        self.test_thread = None  # 添加测试邮件线程

    def load_data(self):
        """延迟加载数据"""
        self.load_templates()
        self.load_last_sender()
        
    def initUI(self):
        """初始化UI"""
        # 设置基本窗口属性
        self.setWindowTitle('BBRHub - 邮件发送工具')
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

    def import_template(self):
        """导入模板"""
        try:
            file_name, _ = QFileDialog.getOpenFileName(
                self,
                "导入模板",
                "",
                "所有支持的格式 (*.eml *.html *.htm *.mht *.txt);;EML文件 (*.eml);;HTML文件 (*.html *.htm);;MHT文件 (*.mht);;文本文件 (*.txt)"
            )
            
            if not file_name:
                return
            
            # 获取文件扩展名
            ext = os.path.splitext(file_name)[1].lower()
            
            try:
                if ext == '.eml':
                    subject, content = self._parse_eml_file(file_name)
                elif ext in ['.html', '.htm']:
                    subject, content = self._parse_html_file(file_name)
                elif ext == '.mht':
                    subject, content = self._parse_mht_file(file_name)
                else:  # .txt
                    with open(file_name, 'r', encoding='utf-8') as f:
                        content = f.read()
                    subject = os.path.splitext(os.path.basename(file_name))[0]
                
                # 确保HTML内容包含基本样式设置
                if ext in ['.html', '.htm', '.eml', '.mht']:
                    content = self._ensure_html_styles(content)
                else:
                    # 将纯文本转换为HTML格式
                    content = (
                        '<!DOCTYPE html>'
                        '<html>'
                        '<head>'
                        '<meta charset="utf-8">'
                        '<meta name="viewport" content="width=device-width, initial-scale=1.0">'
                        '<style>'
                        'body { font-family: Arial, sans-serif; font-size: 14px; line-height: 1.6; color: #333333; margin: 20px; }'
                        'p { margin: 10px 0; }'
                        'a { color: #1890ff; text-decoration: none; }'
                        'a:hover { text-decoration: underline; }'
                        '</style>'
                        '</head>'
                        '<body>'
                        f'<div style="font-size: 14px; line-height: 1.6;">{content.replace(chr(10), "<br>")}</div>'
                        '</body>'
                        '</html>'
                    )
                
                # 弹出对话框让用户输入模板名称
                template_name, ok = QInputDialog.getText(
                    self, 
                    '保存模板',
                    '请输入模板名称:',
                    text=subject
                )
                
                if ok and template_name:
                    # 保存到数据库
                    if self.db.add_template(template_name, content):
                        self.load_templates()  # 重新加载模板列表
                        MessageBox.show(
                            '成功',
                            '模板导入成功',
                            'info',
                            parent=self
                        )
                    else:
                        MessageBox.show(
                            '错误',
                            '模板保存失败',
                            'error',
                            parent=self
                        )
                    
            except Exception as e:
                self.logger.error(f"解析模板文件失败: {str(e)}")
                MessageBox.show(
                    '错误',
                    f'解析模板文件失败: {str(e)}',
                    'error',
                    parent=self
                )
            
        except Exception as e:
            self.logger.error(f"导入模板失败: {str(e)}")
            MessageBox.show(
                '错误',
                f'导入模板失败: {str(e)}',
                'error',
                parent=self
            )

    def _ensure_html_styles(self, content):
        """确保HTML内容包含必要的样式设置"""
        try:
            soup = BeautifulSoup(content, 'html.parser')
            
            # 添加DOCTYPE声明
            if not str(soup).startswith('<!DOCTYPE'):
                content = '<!DOCTYPE html>\n' + str(soup)
                soup = BeautifulSoup(content, 'html.parser')
            
            # 如果没有html标签，添加一个
            if not soup.html:
                new_html = soup.new_tag('html')
                new_html.append(soup)
                soup = BeautifulSoup(str(new_html), 'html.parser')
            
            # 如果没有head标签，添加一个
            if not soup.head:
                head = soup.new_tag('head')
                soup.html.insert(0, head)
            
            # 添加meta标签
            if not soup.head.find('meta', attrs={'charset': True}):
                meta_charset = soup.new_tag('meta', charset='utf-8')
                soup.head.insert(0, meta_charset)
            
            if not soup.head.find('meta', attrs={'name': 'viewport'}):
                meta_viewport = soup.new_tag('meta', attrs={
                    'name': 'viewport',
                    'content': 'width=device-width, initial-scale=1.0'
                })
                soup.head.append(meta_viewport)
            
            # 如果没有body标签，添加一个
            if not soup.body:
                body = soup.new_tag('body')
                if soup.html.contents:
                    body.extend(soup.html.contents[1:])
                soup.html.append(body)
            
            # 添加默认样式
            style = soup.new_tag('style')
            style.string = (
                'body { font-family: Arial, sans-serif; font-size: 14px; line-height: 1.6; color: #333333; margin: 20px; }'
                'p { margin: 0 0 1em 0; }'
                'a { color: #1890ff; text-decoration: none; }'
                'a:hover { text-decoration: underline; }'
                'table { border-collapse: collapse; margin: 1em 0; width: 100%; }'
                'td, th { padding: 8px; border: 1px solid #e8e8e8; }'
            )
            
            # 检查是否已存在style标签
            existing_style = soup.head.find('style')
            if existing_style:
                existing_style.string = style.string + existing_style.string
            else:
                soup.head.append(style)
            
            # 确保所有文本内容都有合适的样式
            for tag in soup.find_all(text=True):
                if tag.parent.name not in ['style', 'script']:
                    if not tag.parent.get('style'):
                        tag.parent['style'] = 'font-size: 14px; line-height: 1.6;'
            
            return str(soup)
            
        except Exception as e:
            self.logger.error(f"处理HTML样式失败: {str(e)}")
            return content

    def _parse_eml_file(self, file_path):
        """解析EML文件"""
        try:
            import email.policy
            
            # 首先检测文件编码
            with open(file_path, 'rb') as f:
                raw_content = f.read()
                detected = chardet.detect(raw_content)
                encoding = detected['encoding'] or 'utf-8'
            
            # 使用检测到的编码读取文件
            with open(file_path, 'r', encoding=encoding, errors='replace') as f:
                msg = message_from_file(f, policy=email.policy.default)
            
            # 获取主题
            subject = msg.get('subject', '')
            if subject.startswith('=?'):  # 处理编码的主题
                subject = str(Header.make_header(Header.decode_header(subject)))
            
            # 获取内容
            content = None
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == 'text/html':
                        try:
                            payload = part.get_payload(decode=True)
                            charset = part.get_content_charset() or 'utf-8'
                            content = payload.decode(charset, errors='replace')
                            break
                        except Exception as e:
                            self.logger.error(f"解析HTML内容失败: {str(e)}")
                            continue
                
                # 如果没有找到HTML内容，尝试获取纯文本
                if not content:
                    for part in msg.walk():
                        if part.get_content_type() == 'text/plain':
                            try:
                                payload = part.get_payload(decode=True)
                                charset = part.get_content_charset() or 'utf-8'
                                content = payload.decode(charset, errors='replace')
                                # 将纯文本转换为HTML格式
                                content = (
                                    '<html>'
                                    '<head>'
                                    '<style>'
                                    'body {'
                                    '    font-family: Arial, sans-serif;'
                                    '    font-size: 14px;'
                                    '    line-height: 1.6;'
                                    '    color: #333333;'
                                    '    margin: 20px;'
                                    '}'
                                    '</style>'
                                    '</head>'
                                    '<body>'
                                    f'<div>{content.replace(chr(10), "<br>")}</div>'
                                    '</body>'
                                    '</html>'
                                )
                                break
                            except Exception as e:
                                self.logger.error(f"解析纯文本内容失败: {str(e)}")
                                continue
            else:
                try:
                    payload = msg.get_payload(decode=True)
                    charset = msg.get_content_charset() or 'utf-8'
                    content = payload.decode(charset, errors='replace')
                except Exception as e:
                    self.logger.error(f"解析非多部分内容失败: {str(e)}")
                    content = msg.get_payload()
            
            if not content:
                raise Exception("无法解析邮件内容")
            
            # 如果内容是纯文本，转换为HTML
            if msg.get_content_type() == 'text/plain':
                content = (
                    '<html>'
                    '<head>'
                    '<style>'
                    'body {'
                    '    font-family: Arial, sans-serif;'
                    '    font-size: 14px;'
                    '    line-height: 1.6;'
                    '    color: #333333;'
                    '    margin: 20px;'
                    '}'
                    '</style>'
                    '</head>'
                    '<body>'
                    f'<div>{content.replace(chr(10), "<br>")}</div>'
                    '</body>'
                    '</html>'
                )
            
            return subject, content
            
        except Exception as e:
            self.logger.error(f"解析EML文件失败: {str(e)}")
            raise Exception(f"解析EML文件失败: {str(e)}")

    def _parse_html_file(self, file_path):
        """解析HTML文件
        
        Args:
            file_path: HTML文件路径
            
        Returns:
            tuple: (主题, 内容)
        """
        from bs4 import BeautifulSoup
        
        with open(file_path, 'r', encoding='utf-8') as f:
            soup = BeautifulSoup(f.read(), 'html.parser')
        
        # 尝试从title标签获取主题
        title = soup.title.string if soup.title else ''
        if not title:
            title = os.path.splitext(os.path.basename(file_path))[0]
        
        # 获取完整HTML内容
        content = str(soup)
        
        return title, content

    def _parse_mht_file(self, file_path):
        """解析MHT文件
        
        Args:
            file_path: MHT文件路径
            
        Returns:
            tuple: (主题, 内容)
        """
        import email.policy
        
        with open(file_path, 'r', encoding='utf-8') as f:
            msg = message_from_file(f, policy=email.policy.default)
        
        # 获取主题
        subject = msg.get('subject', '')
        if subject.startswith('=?'):
            subject = str(Header.make_header(Header.decode_header(subject)))
        if not subject:
            subject = os.path.splitext(os.path.basename(file_path))[0]
        
        # 获取HTML内容
        content = ''
        for part in msg.walk():
            if part.get_content_type() == 'text/html':
                content = part.get_payload(decode=True).decode(part.get_content_charset() or 'utf-8')
                break
        
        return subject, content

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

    def update_test_button(self):
        """更新发送测试按钮的倒计时显示"""
        current_time = time.time()
        remaining_time = int(EmailServer.TEST_EMAIL_INTERVAL - 
                           (current_time - EmailServer.last_test_time))
        
        if remaining_time <= 0:
            self.countdown_timer.stop()
            self.test_btn.setText('发送测试')
            self.test_btn.setEnabled(True)
        else:
            if remaining_time >= 60:
                minutes = remaining_time // 60
                seconds = remaining_time % 60
                time_str = f"{minutes}分{seconds}秒" if seconds else f"{minutes}分钟"
            else:
                time_str = f"{remaining_time}秒"
            self.test_btn.setText(f'发送测试({time_str})')
            self.test_btn.setEnabled(False)

    def send_test_email(self):
        """发送测试邮件到发件人邮箱"""
        # 检查时间间隔
        current_time = time.time()
        if current_time - EmailServer.last_test_time < EmailServer.TEST_EMAIL_INTERVAL:
            remaining_time = int(EmailServer.TEST_EMAIL_INTERVAL - 
                               (current_time - EmailServer.last_test_time))
            if remaining_time >= 60:
                minutes = remaining_time // 60
                seconds = remaining_time % 60
                time_str = f"{minutes}分{seconds}秒" if seconds else f"{minutes}分钟"
            else:
                time_str = f"{remaining_time}秒"
            self.test_btn.setText(f'发送测试({time_str})')
            self.test_btn.setEnabled(False)
            self.countdown_timer.start()
            return

        # 检查条件
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

            # 禁用测试按钮
            self.test_btn.setEnabled(False)
            self.status_label.setText('正在发送测试邮件...')

            # 创建并启动测试邮件线程
            self.test_thread = TestEmailThread(
                self.current_sender,
                template,
                self.config.get_test_data(),
                self.attachments
            )
            self.test_thread.finished.connect(self.on_test_email_finished)
            self.test_thread.start()

        except Exception as e:
            self.logger.error(f"测试邮件发送失败: {str(e)}")
            self.status_label.setText(f'发送失败：{str(e)}')
            self.test_btn.setEnabled(True)

    def on_test_email_finished(self, success, message):
        """测试邮件发送完成处理"""
        if success:
            self.status_label.setText('测试邮件发送成功')
            self.logger.info("测试邮件发送成功")
            # 更新最后发送时间并启动倒计时
            EmailServer.last_test_time = time.time()
            self.test_btn.setEnabled(False)
            self.countdown_timer.start()
        else:
            self.status_label.setText(f'测试邮件发送失败: {message}')
            self.logger.error(f"测试邮件发送失败: {message}")
            self.test_btn.setEnabled(True)

    def start_sending(self):
        """开始批量发送邮件"""
        self.logger.info("开始批量发送邮件")
        if not self.check_send_conditions():
            return
        
        try:
            # 获取当前选中的模板
            template_name = self.template_combo.currentText()
            templates = self.db.get_templates()
            template = templates.get(template_name)
            
            # 先更新UI状态，确保界面响应
            self.progress_bar.setMaximum(len(self.df))
            self.progress_bar.setValue(0)
            self.progress_bar.setFormat('准备发送...')
            self.status_label.setText('正在准备发送...')
            
            # 禁用相关控件
            self.sender_combo.setEnabled(False)
            self.template_combo.setEnabled(False)
            self.test_btn.setEnabled(False)
            
            # 更新发送按钮状态
            self.send_btn.setText('停止发送')
            if self.send_btn.receivers(self.send_btn.clicked) > 0:
                self.send_btn.clicked.disconnect()
            self.send_btn.clicked.connect(self.stop_sending)
            
            # 清理之前的线程（如果存在）
            if self.send_thread and self.send_thread.isRunning():
                self.send_thread.stop()
                self.send_thread.wait()
            
            # 使用 QTimer 延迟创建和启动线程，确保UI更新
            QTimer.singleShot(100, lambda: self._start_send_thread(template))
            
        except Exception as e:
            self.logger.error(f"发送失败: {str(e)}")
            self.status_label.setText(f'发送失败：{str(e)}')
            self._reset_ui_state()

    def _start_send_thread(self, template):
        """启动发送线程"""
        try:
            # 创建发送线程
            self.send_thread = SendEmailThread(
                self.current_sender,
                template,
                self.df,
                self.attachments
            )
            
            # 连接信号
            self.send_thread.progress_updated.connect(self.update_send_progress, Qt.QueuedConnection)
            self.send_thread.finished.connect(self.on_send_finished, Qt.QueuedConnection)
            
            # 启动线程
            self.send_thread.start()
            
        except Exception as e:
            self.logger.error(f"启动发送线程失败: {str(e)}")
            self.status_label.setText(f'启动发送失败：{str(e)}')
            self._reset_ui_state()

    def _reset_ui_state(self):
        """重置UI状态"""
        self.sender_combo.setEnabled(True)
        self.template_combo.setEnabled(True)
        self.test_btn.setEnabled(True)
        self.send_btn.setText('开始发送')
        if self.send_btn.receivers(self.send_btn.clicked) > 0:
            self.send_btn.clicked.disconnect()
        self.send_btn.clicked.connect(self.start_sending)

    def stop_sending(self):
        """停止发送"""
        if self.send_thread and self.send_thread.isRunning():
            self.status_label.setText('正在停止发送...')
            self.send_btn.setEnabled(False)
            self.send_thread.stop()

    def update_send_progress(self, current, status, error):
        """更新发送进度"""
        try:
            # 使用 QTimer 延迟更新UI
            def _update():
                try:
                    self.progress_bar.setValue(current)
                    self.progress_bar.setFormat(f'正在发送 - %p% ({current}/{self.progress_bar.maximum()})')
                    
                    # 更新日志表格
                    row = current - 1
                    if 0 <= row < self.log_table.rowCount():
                        status_item = self.log_table.item(row, 2)
                        if status_item:
                            status_item.setText(status)
                            if '成功' in status:
                                status_item.setForeground(QBrush(QColor('#52c41a')))  # 绿色
                            else:
                                status_item.setForeground(QBrush(QColor('#ff4d4f')))  # 红色
                                status_item.setToolTip(error)  # 设置错误提示
                except Exception as e:
                    self.logger.error(f"更新UI失败: {str(e)}")
            
            QTimer.singleShot(0, _update)
            
        except Exception as e:
            self.logger.error(f"更新进度失败: {str(e)}")

    def on_send_finished(self, success, message):
        """发送完成处理"""
        try:
            self._reset_ui_state()
            
            if success:
                self.progress_bar.setFormat('发送完成 - %p% (%v/%m)')
                self.status_label.setText('发送完成')
            else:
                self.progress_bar.setFormat('发送失败 - %p% (%v/%m)')
                self.status_label.setText(f'发送失败: {message}')
            
            # 记录发送批次结果
            batch_id = f"BATCH_{datetime.now().strftime('%Y%m%d%H%M%S')}"
            self.export_batch_result(batch_id)
            
        except Exception as e:
            self.logger.error(f"处理发送完成事件失败: {str(e)}")
            self.status_label.setText(f'处理发送完成失败：{str(e)}')

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
        dialog = AboutDialog(self)
        dialog.exec_()

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

        # 附件组件 - 移动到右侧
        attachment_layout = QHBoxLayout()
        attachment_label = QLabel('附件:')
        attachment_label.setFixedWidth(80)

        # 显示附件数量
        self.attachment_count = QLabel('无附件')
        self.attachment_count.setStyleSheet("""
            QLabel {
                color: #666666;
                padding: 0 10px;
            }
        """)

        # 查看附件按钮
        view_attachment_btn = QPushButton('管理附件')
        view_attachment_btn.setFixedWidth(80)
        view_attachment_btn.setStyleSheet("""
            QPushButton {
                background-color: #1890ff;
                color: white;
                border: none;
                padding: 3px 8px;
                border-radius: 2px;
                font-size: 12px;
                min-height: 24px;
            }
            QPushButton:hover {
                background-color: #40a9ff;
            }
        """)
        view_attachment_btn.clicked.connect(self.manage_attachments)

        # 组装附件布局
        attachment_layout.addWidget(attachment_label)
        attachment_layout.addWidget(self.attachment_count)
        attachment_layout.addWidget(view_attachment_btn)
        attachment_layout.addStretch()

        # 预览内容
        content_label = QLabel('邮件内容:')
        content_label.setFixedWidth(80)
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

        # 添加到预览布局
        preview_layout.addLayout(title_layout)
        preview_layout.addLayout(attachment_layout)  # 直接添加附件布局
        preview_layout.addWidget(content_label)
        preview_layout.addWidget(self.preview_content)

        # 从左侧布局中移除附件组件
        left_layout.addWidget(sender_group)
        left_layout.addWidget(template_group)
        left_layout.addWidget(progress_group)
        left_layout.addWidget(button_group)
        left_layout.addWidget(self.status_label)
        left_layout.addWidget(log_group)

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

        # 初始化附件列表
        self.attachments = []

    def manage_attachments(self):
        """管理附件"""
        dialog = AttachmentDialog(self, self.attachments)
        if dialog.exec_() == QDialog.Accepted:
            self.attachments = dialog.attachments
            # 更新附件数量显示
            count = len(self.attachments)
            if count > 0:
                total_size = sum(os.path.getsize(f) for f in self.attachments)
                size_str = self.format_file_size(total_size)
                self.attachment_count.setText(f'{count}个附件 ({size_str})')
            else:
                self.attachment_count.setText('无附件')

    def format_file_size(self, size):
        """格式化文件大小显示"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024.0:
                return f"{size:.1f}{unit}"
            size /= 1024.0
        return f"{size:.1f}GB"

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

class TestEmailThread(QThread):
    """测试邮件发送线程"""
    finished = pyqtSignal(bool, str)  # 成功标志, 消息
    
    def __init__(self, sender, template, test_data, attachments=None):
        super().__init__()
        self.sender = sender
        self.template = template
        self.test_data = test_data
        self.attachments = attachments or []
        
    def run(self):
        try:
            # 添加收件人邮箱到测试数据
            self.test_data['收件人邮箱'] = self.sender['email']
            
            # 替换模板中的变量
            content = self.template['content']
            title = self.template['title']
            
            for key, value in self.test_data.items():
                placeholder = '{' + key + '}'
                title = title.replace(placeholder, str(value))
                content = content.replace(placeholder, str(value))
            
            # 创建邮件服务器连接
            server = EmailServer(
                self.sender['email'],
                self.sender['password'],
                self.sender.get('server_type')
            )
            
            # 发送邮件
            success, message = server.send_email(
                self.test_data['收件人邮箱'],
                title,
                content,
                self.attachments
            )
            
            server.close()
            self.finished.emit(success, message)
            
        except Exception as e:
            self.finished.emit(False, str(e))

class SendEmailThread(QThread):
    """邮件发送线程"""
    progress_updated = pyqtSignal(int, str, str)  # 进度, 状态, 错误信息
    finished = pyqtSignal(bool, str)  # 是否成功, 消息
    
    def __init__(self, sender, template, df, attachments=None):
        super().__init__()
        self.sender = sender
        self.template = template
        self.df = df
        self.attachments = attachments or []
        self.is_running = True
        self.server = None
        self.task_queue = Queue()
        self.result_queue = Queue()
        
    def run(self):
        """运行发送任务"""
        try:
            # 创建工作线程
            worker_thread = threading.Thread(target=self._send_worker)
            worker_thread.daemon = True
            worker_thread.start()
            
            # 添加任务到队列
            total = len(self.df)
            for index, row in self.df.iterrows():
                if not self.is_running:
                    break
                self.task_queue.put((index, row))
            
            # 等待所有任务完成或停止
            while self.is_running:
                if self.task_queue.empty() and not worker_thread.is_alive():
                    break
                QThread.msleep(100)
            
            # 处理结果
            success = True
            error_msg = ""
            while not self.result_queue.empty():
                result = self.result_queue.get()
                if not result[0]:  # 如果有任何失败
                    success = False
                    error_msg = result[1]
                    break
            
            if self.is_running:
                self.finished.emit(success, error_msg or '发送完成')
            else:
                self.finished.emit(False, '用户停止发送')
                
        except Exception as e:
            self.finished.emit(False, str(e))
    
    def _send_worker(self):
        """邮件发送工作线程"""
        try:
            # 创建邮件服务器连接
            self.server = EmailServer(
                self.sender['email'],
                self.sender['password'],
                self.sender.get('server_type')
            )
            
            while self.is_running:
                try:
                    # 获取任务，设置超时以便检查停止标志
                    index, row = self.task_queue.get_nowait()  # 改用 get_nowait
                except Queue.Empty:
                    # 如果队列为空，说明任务完成
                    break
                
                try:
                    # 替换变量
                    content = self.template['content']
                    title = self.template['title']
                    data_dict = row.to_dict()
                    
                    for key, value in data_dict.items():
                        placeholder = '{' + key + '}'
                        title = title.replace(placeholder, str(value))
                        content = content.replace(placeholder, str(value))
                    
                    # 发送邮件
                    success, message = self.server.send_email(
                        row['收件人邮箱'],
                        title,
                        content,
                        self.attachments
                    )
                    
                    # 发送进度信号
                    status = '发送成功' if success else f'发送失败: {message}'
                    self.progress_updated.emit(index + 1, status, message if not success else '')
                    
                    # 记录结果
                    self.result_queue.put((success, message))
                    
                    # 标记任务完成
                    self.task_queue.task_done()
                    
                    # 添加短暂延时
                    QThread.msleep(100)
                    
                except Exception as e:
                    self.progress_updated.emit(index + 1, f'发送失败: {str(e)}', str(e))
                    self.result_queue.put((False, str(e)))
                    self.task_queue.task_done()
                    QThread.msleep(200)
            
        finally:
            if self.server:
                try:
                    self.server.close()
                except:
                    pass
    
    def stop(self):
        """停止发送"""
        self.is_running = False
        # 清空任务队列
        while not self.task_queue.empty():
            try:
                self.task_queue.get_nowait()
                self.task_queue.task_done()
            except:
                pass
        # 关闭服务器连接
        if self.server:
            try:
                self.server.close()
            except:
                pass

class AttachmentDialog(QDialog):
    """附件管理对话框"""
    def __init__(self, parent=None, attachments=None):
        super().__init__(parent)
        self.attachments = attachments.copy() if attachments else []  # 创建副本
        self.setWindowTitle('附件管理')
        self.setMinimumWidth(500)
        # 禁用关闭按钮
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowCloseButtonHint)
        self.initUI()
        
    def initUI(self):
        layout = QVBoxLayout()
        
        # 附件列表
        self.attachment_list = QListWidget()
        self.attachment_list.setStyleSheet("""
            QListWidget {
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 5px;
                background-color: #fafafa;
            }
            QListWidget::item {
                padding: 8px;
                border-radius: 4px;
                margin: 2px 0;
            }
            QListWidget::item:hover {
                background-color: #f0f0f0;
            }
            QListWidget::item:selected {
                background-color: #e6f7ff;
                color: #1890ff;
            }
        """)
        
        # 加载现有附件
        for file_path in self.attachments:
            size = os.path.getsize(file_path)
            size_str = self.format_file_size(size)
            item = QListWidgetItem(f"{os.path.basename(file_path)} ({size_str})")
            item.setToolTip(file_path)
            self.attachment_list.addItem(item)
        
        # 按钮布局
        btn_layout = QHBoxLayout()
        add_btn = QPushButton('+ 添加')
        add_btn.setFixedWidth(60)  # 减小按钮宽度
        add_btn.setStyleSheet("""
            QPushButton {
                background-color: #1890ff;
                color: white;
                border: none;
                padding: 4px 8px;
                border-radius: 2px;
                font-size: 12px;
                min-height: 24px;
            }
            QPushButton:hover {
                background-color: #40a9ff;
            }
        """)
        add_btn.clicked.connect(self.add_attachment)
        
        remove_btn = QPushButton('- 删除')
        remove_btn.setFixedWidth(60)  # 减小按钮宽度
        remove_btn.setStyleSheet("""
            QPushButton {
                background-color: #ff4d4f;
                color: white;
                border: none;
                padding: 4px 8px;
                border-radius: 2px;
                font-size: 12px;
                min-height: 24px;
            }
            QPushButton:hover {
                background-color: #ff7875;
            }
        """)
        remove_btn.clicked.connect(self.remove_attachment)
        
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(remove_btn)
        btn_layout.addStretch()
        
        # 总大小显示
        self.size_label = QLabel()
        self.update_total_size()
        
        # 确认和取消按钮
        dialog_buttons = QHBoxLayout()
        
        confirm_btn = QPushButton('确定')
        confirm_btn.setFixedWidth(60)  # 减小按钮宽度
        confirm_btn.setStyleSheet("""
            QPushButton {
                background-color: #1890ff;
                color: white;
                border: none;
                padding: 4px 8px;
                border-radius: 2px;
                font-size: 12px;
                min-height: 24px;
            }
            QPushButton:hover {
                background-color: #40a9ff;
            }
        """)
        confirm_btn.clicked.connect(self.accept)
        
        cancel_btn = QPushButton('取消')
        cancel_btn.setFixedWidth(60)  # 减小按钮宽度
        cancel_btn.setStyleSheet("""
            QPushButton {
                background-color: #f5f5f5;
                color: #333;
                border: 1px solid #d9d9d9;
                padding: 4px 8px;
                border-radius: 2px;
                font-size: 12px;
                min-height: 24px;
            }
            QPushButton:hover {
                background-color: #fafafa;
                border-color: #40a9ff;
                color: #40a9ff;
            }
        """)
        cancel_btn.clicked.connect(self.reject)
        
        dialog_buttons.addStretch()
        dialog_buttons.addWidget(confirm_btn)
        dialog_buttons.addWidget(cancel_btn)
        
        # 添加到主布局
        layout.addWidget(self.attachment_list)
        layout.addLayout(btn_layout)
        layout.addWidget(self.size_label)
        layout.addLayout(dialog_buttons)  # 添加确认取消按钮
        
        self.setLayout(layout)
        
    def format_file_size(self, size):
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024.0:
                return f"{size:.1f}{unit}"
            size /= 1024.0
        return f"{size:.1f}GB"
        
    def update_total_size(self):
        if self.attachments:
            total = sum(os.path.getsize(f) for f in self.attachments)
            self.size_label.setText(f'总大小: {self.format_file_size(total)}')
        else:
            self.size_label.setText('没有附件')
            
    def add_attachment(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择附件",
            "",
            "所有文件 (*.*)"
        )
        
        if files:
            total_size = sum(os.path.getsize(f) for f in self.attachments)
            
            for file_path in files:
                file_size = os.path.getsize(file_path)
                
                if file_size > EmailServer.MAX_ATTACHMENT_SIZE:
                    MessageBox.show(
                        '错误',
                        f'附件 {os.path.basename(file_path)} 超过大小限制(25MB)',
                        'critical',
                        parent=self
                    )
                    continue
                    
                if total_size + file_size > EmailServer.MAX_ATTACHMENT_SIZE:
                    MessageBox.show(
                        '错误',
                        '附件总大小超过限制(25MB)',
                        'critical',
                        parent=self
                    )
                    break
                    
                if file_path not in self.attachments:
                    self.attachments.append(file_path)
                    size_str = self.format_file_size(file_size)
                    item = QListWidgetItem(f"{os.path.basename(file_path)} ({size_str})")
                    item.setToolTip(file_path)
                    self.attachment_list.addItem(item)
                    total_size += file_size
            
            self.update_total_size()
            
    def remove_attachment(self):
        current_row = self.attachment_list.currentRow()
        if current_row >= 0:
            self.attachment_list.takeItem(current_row)
            self.attachments.pop(current_row)
            self.update_total_size()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    # 修改图标加载方式
    app.setWindowIcon(QIcon(get_resource_path('imgs/logo.ico')))
    ex = EmailSender()
    ex.show()
    sys.exit(app.exec_()) 