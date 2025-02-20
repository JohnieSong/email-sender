from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                            QLineEdit, QTextEdit, QPushButton, QMessageBox,
                            QListWidget, QSplitter, QWidget, QToolButton, QMenu, QTextBrowser, QFileDialog, QInputDialog)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont
import json
import os
from message_box import MessageBox
from styles import MODERN_STYLE  # 添加到导入部分
from database import Database
from config import Config
from email import message_from_file
import email.policy
from email.header import Header
from bs4 import BeautifulSoup

class TemplateDialog(QDialog):
    # 添加信号
    template_updated = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.config = Config()  # 先创建 config 实例
        self.db = Database(self.config)  # 传入 config 实例
        self.setWindowTitle('模板管理')
        self.setStyleSheet(MODERN_STYLE)
        self.setMinimumWidth(1000)
        self.setMinimumHeight(800)
        
        self.current_template = None
        self.last_focused = None
        self.variables = []  # 初始化变量列表为空
        self.templates = {}  # 初始化模板字典为空
        
        # 初始化UI
        self.initUI()
        
        # 加载数据并更新UI
        self.variables = self.load_variables()
        self.templates = self.load_templates()
        self.load_template_list()

    def initUI(self):
        self.setWindowTitle('模板管理')
        self.setGeometry(200, 200, 1000, 600)
        
        layout = QHBoxLayout()
        
        # 左侧模板列表区域
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        
        # 模板列表
        self.template_list = QListWidget()
        self.template_list.setSelectionMode(QListWidget.ExtendedSelection)  # 允许多选
        self.template_list.itemSelectionChanged.connect(self.on_selection_changed)
        
        # 模板操作按钮
        btn_layout = QHBoxLayout()
        
        # 创建按钮
        add_btn = QPushButton('添加')
        import_btn = QPushButton('导入')
        export_btn = QPushButton('导出')  # 将修改按钮改为导出按钮
        delete_btn = QPushButton('删除')
        
        # 连接按钮事件
        add_btn.clicked.connect(self.add_template)
        import_btn.clicked.connect(self.import_template)
        export_btn.clicked.connect(self.export_template)  # 连接导出事件
        delete_btn.clicked.connect(self.delete_template)
        
        # 将按钮添加到布局中
        btn_layout.addStretch()
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(import_btn)
        btn_layout.addWidget(export_btn)  # 添加导出按钮
        btn_layout.addWidget(delete_btn)
        btn_layout.addStretch()
        
        # 添加组件到左侧布局
        left_layout.addWidget(QLabel('已保存的模板:'))
        left_layout.addWidget(self.template_list)
        left_layout.addLayout(btn_layout)
        
        # 右侧编辑区域
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        
        # 模板名称
        name_layout = QHBoxLayout()
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText('输入模板名称')
        name_layout.addWidget(QLabel('模板名称:'))
        name_layout.addWidget(self.name_input)
        
        # 邮件标题
        title_layout = QHBoxLayout()
        self.title_input = QLineEdit()
        self.title_input.setPlaceholderText('输入邮件标题 (支持使用{姓名}等变量)')
        
        # 更新插入变量按钮的样式和行为
        self.insert_var_btn = QPushButton('插入变量 ▼')
        self.insert_var_btn.setStyleSheet("""
            QPushButton {
                background-color: #1890ff;
                color: white;
                border: none;
                padding: 6px 16px;
                border-radius: 3px;
                min-height: 32px;
                font-size: 14px;
                font-family: "Microsoft YaHei", "Segoe UI";
            }
            QPushButton:hover {
                background-color: #40a9ff;
            }
            QPushButton:pressed {
                background-color: #096dd9;
            }
            QPushButton::menu-indicator {
                width: 0px;
            }
        """)
        
        # 直接连接按钮点击事件到显示菜单方法
        self.insert_var_btn.clicked.connect(self.show_var_menu)
        
        title_layout.addWidget(QLabel('邮件标题:'))
        title_layout.addWidget(self.title_input)
        title_layout.addWidget(self.insert_var_btn)
        
        # 创建邮件内容编辑器
        self.content_edit = QTextEdit()
        self.content_edit.setPlaceholderText('输入邮件内容，支持HTML格式和变量替换')
        self.content_edit.setAcceptRichText(True)
        self.content_edit.setAutoFormatting(QTextEdit.AutoAll)
        
        # 设置文档格式，允许图片和链接
        doc = self.content_edit.document()
        doc.setDefaultStyleSheet("""
            a { color: #0000ff; text-decoration: underline; }
            img { max-width: 100%; }
        """)
        
        # 将组件添加到右侧布局
        right_layout.addLayout(name_layout)
        right_layout.addLayout(title_layout)
        right_layout.addWidget(QLabel('邮件内容:'))
        right_layout.addWidget(self.content_edit)
        
        # 按钮
        button_layout = QHBoxLayout()
        save_btn = QPushButton('保存')
        save_btn.clicked.connect(self.save_template)
        close_btn = QPushButton('关闭')
        close_btn.clicked.connect(self.accept)
        
        button_layout.addWidget(save_btn)
        button_layout.addWidget(close_btn)
        
        # 添加按钮到右侧布局
        right_layout.addLayout(button_layout)
        
        # 使用分割器组合左右两侧
        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setSizes([300, 700])
        
        layout.addWidget(splitter)
        self.setLayout(layout)
        
        # 标题输入框焦点变化
        self.title_input.focusInEvent = lambda e: self.on_focus_change(self.title_input, e)
        
        # 内容编辑器焦点变化
        self.content_edit.focusInEvent = lambda e: self.on_focus_change(self.content_edit, e)

    def load_variables(self):
        """从数据库加载变量列表"""
        variables = self.db.get_variables()
        return [var['name'] for var in variables]  # 返回变量名列表

    def load_templates(self):
        """从数据库加载模板"""
        try:
            return self.db.get_templates()
        except Exception as e:
            MessageBox.show_error(f"加载模板失败: {str(e)}")
            return {}

    def load_template_list(self):
        """加载模板列表"""
        current_template = self.template_list.currentItem()
        current_name = current_template.text() if current_template else None
        
        self.template_list.clear()
        self.template_list.addItems(self.templates.keys())
        
        # 尝试恢复之前选中的模板
        if current_name:
            items = self.template_list.findItems(current_name, Qt.MatchExactly)
            if items:
                self.template_list.setCurrentItem(items[0])

    def save_template(self):
        """保存当前模板"""
        name = self.name_input.text().strip()
        title = self.title_input.text().strip()
        content = self.content_edit.toHtml()
        
        if not all([name, title, content]):
            MessageBox.show('错误', '请填写所有字段', 'warning', parent=self)
            return
            
        try:
            self.db.add_template(name, title, content)
            # 重新加载模板列表
            self.templates = self.load_templates()
            self.load_template_list()
            
            # 选中新保存的模板
            items = self.template_list.findItems(name, Qt.MatchExactly)
            if items:
                self.template_list.setCurrentItem(items[0])
            
            self.template_updated.emit()
            MessageBox.show('成功', f'模板"{name}"已保存', 'info', parent=self)
        except Exception as e:
            MessageBox.show('错误', f'保存失败：{str(e)}', 'critical', parent=self)

    def add_template(self):
        """新建模板"""
        self.current_template = None
        self.name_input.clear()
        self.title_input.clear()
        self.content_edit.clear()
        self.name_input.setFocus()
        
        # 取消选中当前模板
        self.template_list.clearSelection()

    def delete_template(self):
        """删除当前选中的模板"""
        current_item = self.template_list.currentItem()
        if not current_item:
            return
            
        template_name = current_item.text()
        reply = MessageBox.show(
            '确认删除', 
            f'确定要删除模板"{template_name}"吗？',
            'question',
            [('是', QMessageBox.AcceptRole), ('否', QMessageBox.RejectRole)],
            parent=self
        )
        
        if reply == 0:
            if self.db.remove_template(template_name):
                # 重新加载模板列表
                self.templates = self.load_templates()
                self.load_template_list()
                
                # 清空编辑区域
                self.add_template()
                
                self.template_updated.emit()
                MessageBox.show('成功', f'模板"{template_name}"已删除', 'info', parent=self)
            else:
                MessageBox.show('错误', '删除失败', 'critical', parent=self)

    def on_selection_changed(self):
        """处理选择变化"""
        selected_items = self.template_list.selectedItems()
        current_item = self.template_list.currentItem()
        
        if len(selected_items) == 1 and current_item:
            # 单选情况
            template_name = current_item.text()
            self.current_template = template_name
            template = self.templates.get(template_name)
            if template:
                self.name_input.setText(template_name)
                self.title_input.setText(template['title'])
                self.content_edit.setHtml(template['content'])
        else:
            # 多选或未选择情况
            self.current_template = None
            self.name_input.clear()
            self.title_input.clear()
            self.content_edit.clear()

    def on_focus_change(self, widget, event):
        """记录最后获得焦点的编辑器"""
        self.last_focused = widget
        # 调用原始的焦点事件处理
        type(widget).focusInEvent(widget, event)

    def insert_variable_to_focused(self, var):
        """根据焦点插入变量"""
        # 使用保存的焦点或当前焦点
        target_widget = self.last_focused or self.focusWidget()
        
        if target_widget in [self.title_input, self.content_edit]:
            if isinstance(target_widget, QLineEdit):
                # 对于标题输入框
                current_text = target_widget.text()
                current_pos = target_widget.cursorPosition()
                new_text = current_text[:current_pos] + f"{{{var}}}" + current_text[current_pos:]
                target_widget.setText(new_text)
                # 将光标移动到插入的变量后面
                target_widget.setCursorPosition(current_pos + len(var) + 2)
            else:
                # 对于内容编辑器
                cursor = target_widget.textCursor()
                cursor.insertText(f"{{{var}}}")
                target_widget.setFocus()  # 重新获得焦点

    def open_link(self, url):
        """打开链接"""
        import webbrowser
        webbrowser.open(url.toString())

    def import_template(self):
        """导入模板"""
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "导入模板",
            "",
            "所有支持的格式 (*.eml *.html *.htm *.mht *.txt *.json);;EML文件 (*.eml);;HTML文件 (*.html *.htm);;MHT文件 (*.mht);;文本文件 (*.txt);;JSON文件 (*.json)"
        )
        
        if not file_name:
            return
        
        try:
            # 获取文件扩展名
            ext = os.path.splitext(file_name)[1].lower()
            
            if ext == '.json':
                self._import_json_template(file_name)
            else:
                self._import_mail_template(file_name, ext)
            
        except Exception as e:
            MessageBox.show(
                '错误',
                f'导入失败：{str(e)}',
                'critical',
                parent=self
            )

    def _import_json_template(self, file_path):
        """导入JSON格式的模板"""
        with open(file_path, 'r', encoding='utf-8') as f:
            templates = json.load(f)
            
        if not isinstance(templates, dict):
            raise ValueError('无效的模板文件格式')
            
        # 检查并导入模板
        for name, template in templates.items():
            if not isinstance(template, dict) or 'title' not in template or 'content' not in template:
                raise ValueError(f'模板"{name}"格式无效')
                
            self.db.add_template(name, template['title'], template['content'])
        
        # 重新加载模板列表
        self.templates = self.load_templates()
        self.load_template_list()
        
        # 发送模板更新信号
        self.template_updated.emit()
        
        MessageBox.show(
            '成功',
            f'已导入 {len(templates)} 个模板',
            'info',
            parent=self
        )

    def _import_mail_template(self, file_path, ext):
        """导入邮件格式的模板"""
        if ext == '.eml':
            subject, content = self._parse_eml_file(file_path)
        elif ext in ['.html', '.htm']:
            subject, content = self._parse_html_file(file_path)
        elif ext == '.mht':
            subject, content = self._parse_mht_file(file_path)
        else:  # .txt
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            subject = os.path.splitext(os.path.basename(file_path))[0]
        
        # 使用文件名作为默认模板名称
        template_name = os.path.splitext(os.path.basename(file_path))[0]
        
        # 弹出对话框让用户确认或修改模板名称
        template_name, ok = QInputDialog.getText(
            self,
            '保存模板',
            '请输入模板名称:',
            text=template_name
        )
        
        if ok and template_name:
            self.db.add_template(template_name, subject, content)
            self.templates = self.load_templates()
            self.load_template_list()
            
            # 发送模板更新信号
            self.template_updated.emit()
            
            MessageBox.show(
                '成功',
                '模板导入成功',
                'info',
                parent=self
            )

    def show_var_menu(self):
        """显示变量菜单"""
        # 获取变量列表
        variables = self.load_variables()
        
        # 如果没有变量，显示提示信息
        if not variables:
            MessageBox.show(
                '提示', 
                '当前没有可用变量。\n您可以通过以下方式添加变量：\n'
                '1. 导入Excel文件，表头将自动添加为变量\n'
                '2. 在"工具-变量管理"中手动添加变量',
                'info',
                parent=self
            )
            return
        
        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                background-color: white;
                border: 1px solid #e8e8e8;
                border-radius: 3px;
                padding: 8px;
                min-width: 150px;
            }
            QMenu::item {
                padding: 8px 20px;
                border-radius: 3px;
                font-size: 14px;
                font-family: "Microsoft YaHei", "Segoe UI";
                color: #333333;
                margin: 2px 4px;
            }
            QMenu::item:selected {
                background-color: #e6f7ff;
                color: #1890ff;
            }
            QMenu::separator {
                height: 1px;
                background-color: #e8e8e8;
                margin: 4px 0px;
            }
            QMenu::section {
                padding: 4px 20px;
                color: #666666;
                font-weight: bold;
                background-color: #f5f5f5;
            }
        """)
        
        # 添加常用变量组
        common_vars = ['姓名', '收件人邮箱']
        common_group = []
        other_group = []
        
        for var in variables:
            if var in common_vars:
                common_group.append(var)
            else:
                other_group.append(var)
        
        # 添加常用变量
        if common_group:
            menu.addSection("常用变量")
            for var in common_group:
                action = menu.addAction(f'{{{var}}}')
                action.triggered.connect(lambda checked, v=var: self.insert_variable_to_focused(v))
        
        # 添加分隔线
        if common_group and other_group:
            menu.addSeparator()
        
        # 添加其他变量
        if other_group:
            menu.addSection("其他变量")
            for var in other_group:
                action = menu.addAction(f'{{{var}}}')
                action.triggered.connect(lambda checked, v=var: self.insert_variable_to_focused(v))
        
        # 添加底部操作项
        menu.addSeparator()
        manage_action = menu.addAction("管理变量...")
        manage_action.triggered.connect(self.open_variable_manager)
        
        # 在按钮位置显示菜单
        menu.exec_(self.insert_var_btn.mapToGlobal(self.insert_var_btn.rect().bottomLeft()))

    def open_variable_manager(self):
        """打开变量管理器"""
        from variable_dialog import VariableDialog
        dialog = VariableDialog(self)
        dialog.variable_updated.connect(self.on_variables_updated)
        dialog.exec_()

    def on_variables_updated(self):
        """当变量更新时刷新变量列表"""
        self.variables = self.load_variables()

    def save_current_focus(self):
        """保存当前焦点"""
        current = self.focusWidget()
        if current in [self.title_input, self.content_edit]:
            self.last_focused = current 

    def export_template(self):
        """导出选中的模板"""
        selected_items = self.template_list.selectedItems()
        if not selected_items:
            MessageBox.show('提示', '请选择要导出的模板', 'warning', parent=self)
            return
        
        # 获取选中的模板数据
        export_templates = {}
        for item in selected_items:
            template_name = item.text()
            template = self.templates.get(template_name)
            if template:
                export_templates[template_name] = template
        
        if not export_templates:
            return
        
        # 根据选中数量决定可用的导出格式
        if len(export_templates) == 1:
            # 单个模板可以导出所有格式
            filter_str = "JSON文件 (*.json);;HTML文件 (*.html);;HTM文件 (*.htm);;MHT文件 (*.mht);;EML文件 (*.eml);;文本文件 (*.txt)"
        else:
            # 多个模板只能导出部分格式
            filter_str = "JSON文件 (*.json);;HTML文件 (*.html);;HTM文件 (*.htm);;文本文件 (*.txt)"
        
        # 选择导出格式和文件名
        default_name = list(export_templates.keys())[0] if len(export_templates) == 1 else 'templates'
        file_name, selected_filter = QFileDialog.getSaveFileName(
            self,
            "导出模板",
            default_name,
            filter_str
        )
        
        if not file_name:
            return
        
        try:
            # 根据选择的格式导出
            if selected_filter == 'JSON文件 (*.json)':
                self._export_json_template(file_name, export_templates)
            elif selected_filter in ['HTML文件 (*.html)', 'HTM文件 (*.htm)']:
                self._export_html_template(file_name, export_templates)
            elif selected_filter == 'MHT文件 (*.mht)':
                self._export_mht_template(file_name, export_templates)
            elif selected_filter == 'EML文件 (*.eml)':
                self._export_eml_template(file_name, export_templates)
            else:  # 文本文件
                self._export_text_template(file_name, export_templates)
            
            MessageBox.show(
                '成功',
                f'已导出 {len(export_templates)} 个模板',
                'info',
                parent=self
            )
            
        except Exception as e:
            MessageBox.show(
                '错误',
                f'导出失败：{str(e)}',
                'critical',
                parent=self
            )

    def _export_json_template(self, file_path, templates):
        """导出为JSON格式"""
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(templates, f, ensure_ascii=False, indent=4)

    def _export_html_template(self, file_path, templates):
        """导出为HTML格式"""
        html_content = []
        for name, template in templates.items():
            html_content.append(f'<h1>{template["title"]}</h1>')
            html_content.append(template["content"])
            html_content.append('<hr>')
        
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(html_content))

    def _export_eml_template(self, file_path, templates):
        """导出为EML格式"""
        from email.mime.text import MIMEText
        from email.mime.multipart import MIMEMultipart
        
        for name, template in templates.items():
            msg = MIMEMultipart('alternative')
            msg['Subject'] = template['title']
            msg['From'] = "template@example.com"
            msg['To'] = "recipient@example.com"
            
            html_part = MIMEText(template['content'], 'html', 'utf-8')
            msg.attach(html_part)
            
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(msg.as_string())
            break  # EML格式只导出第一个模板

    def _export_text_template(self, file_path, templates):
        """导出为文本格式"""
        text_content = []
        for name, template in templates.items():
            text_content.append(f'标题: {template["title"]}')
            text_content.append('内容:')
            text_content.append(template["content"])
            text_content.append('-' * 50)
        
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(text_content))

    def _export_mht_template(self, file_path, templates):
        """导出为MHT格式"""
        from email.mime.text import MIMEText
        from email.mime.multipart import MIMEMultipart
        from email.generator import Generator
        import datetime
        
        for name, template in templates.items():
            msg = MIMEMultipart('related')
            msg['Subject'] = template['title']
            msg['From'] = "template@example.com"
            msg['To'] = "recipient@example.com"
            msg['Date'] = datetime.datetime.now().strftime("%a, %d %b %Y %H:%M:%S +0000")
            msg['MIME-Version'] = '1.0'
            msg['Content-Type'] = 'multipart/related'
            
            html_part = MIMEText(template['content'], 'html', 'utf-8')
            msg.attach(html_part)
            
            with open(file_path, 'w', encoding='utf-8') as f:
                gen = Generator(f)
                gen.flatten(msg)
            break  # MHT格式只导出第一个模板

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
                # 从 config.ini 获取测试数据
                test_data = self.config.get_test_data()
                if not test_data:  # 如果没有测试数据，创建默认值
                    test_data = {
                        'test_sender': 'test@example.com',
                        'test_recipient': 'recipient@example.com',
                        'test_subject': '测试主题',
                        'test_content': '测试内容'
                    }
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
            print(f"加载模板失败: {str(e)}")

    def edit_template(self):
        """编辑当前选中的模板"""
        current_item = self.template_list.currentItem()
        if not current_item:
            MessageBox.show('提示', '请先选择要修改的模板', 'warning', parent=self)
            return
        
        template_name = current_item.text()
        template = self.templates.get(template_name)
        if template:
            self.current_template = template_name
            self.name_input.setText(template_name)
            self.title_input.setText(template['title'])
            self.content_edit.setHtml(template['content'])
            self.name_input.setFocus() 

    def _parse_eml_file(self, file_path):
        """解析EML文件"""
        with open(file_path, 'r', encoding='utf-8') as f:
            msg = message_from_file(f, policy=email.policy.default)
        
        # 获取主题
        subject = msg.get('subject', '')
        if subject.startswith('=?'):  # 处理编码的主题
            subject = str(Header.make_header(Header.decode_header(subject)))
        
        # 获取内容
        content = ''
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == 'text/html':
                    content = part.get_payload(decode=True).decode(part.get_content_charset() or 'utf-8')
                    break
            if not content:  # 如果没有HTML内容,尝试获取纯文本
                for part in msg.walk():
                    if part.get_content_type() == 'text/plain':
                        content = part.get_payload(decode=True).decode(part.get_content_charset() or 'utf-8')
                        break
        else:
            content = msg.get_payload(decode=True).decode(msg.get_content_charset() or 'utf-8')
        
        return subject, content

    def _parse_html_file(self, file_path):
        """解析HTML文件"""
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
        """解析MHT文件"""
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