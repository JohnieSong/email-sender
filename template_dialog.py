from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                            QLineEdit, QTextEdit, QPushButton, QMessageBox,
                            QListWidget, QSplitter, QWidget, QToolButton, QMenu, QTextBrowser, QFileDialog)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont
import json
import os
from message_box import MessageBox
from styles import MODERN_STYLE  # 添加到导入部分
from database import Database
from config import Config

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
        list_buttons = QHBoxLayout()
        new_btn = QPushButton('新建')
        new_btn.clicked.connect(self.new_template)
        import_btn = QPushButton('导入')
        import_btn.clicked.connect(self.import_template)
        export_btn = QPushButton('导出')
        export_btn.clicked.connect(self.export_template)
        delete_btn = QPushButton('删除')
        delete_btn.clicked.connect(self.delete_template)
        list_buttons.addWidget(new_btn)
        list_buttons.addWidget(import_btn)
        list_buttons.addWidget(export_btn)
        list_buttons.addWidget(delete_btn)
        
        # 添加组件到左侧布局
        left_layout.addWidget(QLabel('已保存的模板:'))
        left_layout.addWidget(self.template_list)
        left_layout.addLayout(list_buttons)
        
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

    def new_template(self):
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
                self.new_template()
                
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
            "选择模板文件",
            "",
            "JSON Files (*.json);;All Files (*)"
        )
        
        if not file_name:
            return
        
        try:
            with open(file_name, 'r', encoding='utf-8') as f:
                imported_templates = json.load(f)
                
            if not isinstance(imported_templates, dict):
                MessageBox.show('错误', '无效的模板文件格式', 'warning', parent=self)
                return
                
            # 检查模板格式
            for name, template in imported_templates.items():
                if not isinstance(template, dict) or 'title' not in template or 'content' not in template:
                    MessageBox.show('错误', f'模板"{name}"格式无效', 'warning', parent=self)
                    return
            
            # 确认是否覆盖已存在的模板
            existing_templates = set(self.templates.keys()) & set(imported_templates.keys())
            if existing_templates:
                reply = MessageBox.show(
                    '确认导入',
                    f'以下模板已存在，是否覆盖？\n{", ".join(existing_templates)}',
                    'question',
                    [('是', QMessageBox.AcceptRole), ('否', QMessageBox.RejectRole)],
                    parent=self
                )
                if reply != 0:  # 用户选择不覆盖
                    return
            
            # 导入模板
            for name, template in imported_templates.items():
                self.db.add_template(name, template['title'], template['content'])
            
            # 重新加载模板列表
            self.templates = self.load_templates()
            self.load_template_list()
            
            # 选中第一个导入的模板
            if imported_templates:
                first_template = list(imported_templates.keys())[0]
                items = self.template_list.findItems(first_template, Qt.MatchExactly)
                if items:
                    self.template_list.setCurrentItem(items[0])
            
            self.template_updated.emit()
            
            MessageBox.show(
                '成功', 
                f'已导入 {len(imported_templates)} 个模板', 
                'info', 
                parent=self
            )
            
        except json.JSONDecodeError:
            MessageBox.show('错误', '无效的JSON文件', 'critical', parent=self)
        except Exception as e:
            MessageBox.show('错误', f'导入失败：{str(e)}', 'critical', parent=self)

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
            MessageBox.show('提示', '请按住 Ctrl 键或按住 Shift 键点击选择要导出的模板', 'warning', parent=self)
            return
        
        # 如果只选择了一个模板，使用模板名作为默认文件名
        if len(selected_items) == 1:
            default_filename = f"{selected_items[0].text()}.json"
        else:
            default_filename = "templates.json"
        
        # 获取保存文件路径
        file_name, _ = QFileDialog.getSaveFileName(
            self,
            "导出模板",
            default_filename,
            "JSON Files (*.json);;All Files (*)"
        )
        
        if not file_name:
            return
        
        try:
            # 创建导出数据
            export_data = {}
            for item in selected_items:
                template_name = item.text()
                template = self.templates.get(template_name)
                if template:
                    export_data[template_name] = {
                        'title': template['title'],
                        'content': template['content']
                    }
            
            # 写入文件
            with open(file_name, 'w', encoding='utf-8') as f:
                json.dump(export_data, f, ensure_ascii=False, indent=4)
            
            # 根据导出的模板数量显示不同的成功消息
            if len(export_data) == 1:
                message = f'模板"{list(export_data.keys())[0]}"已导出到：\n{file_name}'
            else:
                message = f'已成功导出 {len(export_data)} 个模板到：\n{file_name}'
            
            MessageBox.show(
                '成功',
                message,
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