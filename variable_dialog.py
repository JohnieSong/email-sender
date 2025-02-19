from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                            QLineEdit, QPushButton, QListWidget, QMessageBox)
from PyQt5.QtCore import Qt, pyqtSignal
from message_box import MessageBox
from styles import MODERN_STYLE
from database import Database
from config import Config

class VariableDialog(QDialog):
    variable_updated = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.config = Config()
        self.db = Database(self.config)
        self.setWindowTitle('变量管理')
        self.setStyleSheet(MODERN_STYLE)
        self.setMinimumSize(300, 400)
        self.initUI()
        self.load_variables()
        
    def initUI(self):
        layout = QVBoxLayout()
        layout.setSpacing(10)
        
        # 变量列表
        self.var_list = QListWidget()
        
        # 启用工具提示
        self.var_list.setMouseTracking(True)
        self.var_list.setToolTipDuration(5000)
        
        # 按钮组
        btn_layout = QHBoxLayout()
        add_btn = QPushButton('添加')
        edit_btn = QPushButton('修改')
        delete_btn = QPushButton('删除')
        close_btn = QPushButton('关闭')
        
        add_btn.clicked.connect(self.add_variable)
        edit_btn.clicked.connect(self.edit_variable)
        delete_btn.clicked.connect(self.delete_variable)
        close_btn.clicked.connect(self.accept)
        
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(edit_btn)
        btn_layout.addWidget(delete_btn)
        btn_layout.addStretch()
        btn_layout.addWidget(close_btn)
        
        layout.addWidget(QLabel('变量列表:'))
        layout.addWidget(self.var_list)
        layout.addLayout(btn_layout)
        
        self.setLayout(layout)
        
    def load_variables(self):
        """加载变量列表"""
        self.var_list.clear()
        variables = self.db.get_variables()
        for var in variables:
            self.var_list.addItem(var['name'])
            
    def add_variable(self):
        """添加新变量"""
        dialog = QDialog(self)
        dialog.setWindowTitle('添加变量')
        dialog.setModal(True)
        dialog.setStyleSheet(MODERN_STYLE)
        
        layout = QVBoxLayout()
        
        var_input = QLineEdit()
        var_input.setPlaceholderText('输入变量名称')
        
        btn_layout = QHBoxLayout()
        ok_btn = QPushButton('确定')
        cancel_btn = QPushButton('取消')
        
        ok_btn.clicked.connect(dialog.accept)
        cancel_btn.clicked.connect(dialog.reject)
        
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        
        layout.addWidget(QLabel('变量名称:'))
        layout.addWidget(var_input)
        layout.addLayout(btn_layout)
        
        dialog.setLayout(layout)
        
        if dialog.exec_() == QDialog.Accepted:
            var_name = var_input.text().strip()
            if var_name:
                if var_name not in [self.var_list.item(i).text() for i in range(self.var_list.count())]:
                    if self.db.add_variable(var_name):
                        self.load_variables()
                        self.variable_updated.emit()
                        MessageBox.show('成功', f'已添加变量"{var_name}"', 'info', parent=self)
                    else:
                        MessageBox.show('错误', '添加变量失败', 'warning', parent=self)
                else:
                    MessageBox.show('错误', '变量已存在', 'warning', parent=self)
                    
    def edit_variable(self):
        """编辑变量"""
        current_item = self.var_list.currentItem()
        if not current_item:
            MessageBox.show('提示', '请先选择要编辑的变量', 'warning', parent=self)
            return
            
        old_var = current_item.text()
        
        dialog = QDialog(self)
        dialog.setWindowTitle('编辑变量')
        dialog.setStyleSheet(MODERN_STYLE)
        dialog.setFixedWidth(300)
        
        layout = QVBoxLayout(dialog)
        layout.setSpacing(15)
        
        var_input = QLineEdit(old_var)
        var_input.setPlaceholderText('输入变量名')
        
        btn_layout = QHBoxLayout()
        save_btn = QPushButton('保存')
        save_btn.clicked.connect(dialog.accept)
        cancel_btn = QPushButton('取消')
        cancel_btn.clicked.connect(dialog.reject)
        
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        
        layout.addWidget(QLabel('变量名:'))
        layout.addWidget(var_input)
        layout.addLayout(btn_layout)
        
        if dialog.exec_() == QDialog.Accepted:
            new_var = var_input.text().strip()
            if new_var:
                if new_var != old_var:
                    if new_var not in [self.var_list.item(i).text() for i in range(self.var_list.count())]:
                        try:
                            # 更新数据库中的变量
                            self.db.update_variable(old_var, new_var)
                            
                            # 更新所有模板中的变量
                            templates = self.db.get_templates()
                            for name, template in templates.items():
                                title = template['title'].replace(
                                    f'{{{old_var}}}', f'{{{new_var}}}'
                                )
                                content = template['content'].replace(
                                    f'{{{old_var}}}', f'{{{new_var}}}'
                                )
                                self.db.update_template(name, title, content)
                            
                            self.load_variables()
                            self.variable_updated.emit()
                            MessageBox.show('成功', f'变量已更新', 'info', parent=self)
                            
                        except Exception as e:
                            MessageBox.show('错误', f'保存失败：{str(e)}', 'critical', parent=self)
                    else:
                        MessageBox.show('错误', '变量名已存在', 'warning', parent=self)
                        
    def delete_variable(self):
        """删除变量"""
        current_item = self.var_list.currentItem()
        if not current_item:
            MessageBox.show('提示', '请先选择要删除的变量', 'warning', parent=self)
            return
            
        var_name = current_item.text()
        
        if MessageBox.confirm('确认', f'确定要删除变量"{var_name}"吗？\n删除后所有使用此变量的模板都需要更新。', parent=self):
            try:
                if self.db.remove_variable(var_name):
                    self.load_variables()
                    self.variable_updated.emit()
                    MessageBox.show('成功', f'已删除变量"{var_name}"', 'info', parent=self)
                else:
                    MessageBox.show('错误', '删除变量失败', 'warning', parent=self)
            except Exception as e:
                MessageBox.show('错误', f'删除失败：{str(e)}', 'critical', parent=self) 