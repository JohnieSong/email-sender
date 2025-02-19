from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                            QPushButton, QTableWidget, QTableWidgetItem,
                            QDateEdit, QComboBox, QHeaderView, QWidget)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QColor, QBrush
import pandas as pd
from database import Database
from config import Config
from message_box import MessageBox

class LogDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.config = Config()
        self.db = Database(self.config)
        self.setWindowTitle('发送历史记录')
        self.setMinimumWidth(1000)
        self.setMinimumHeight(600)
        self.initUI()
        
    def initUI(self):
        layout = QVBoxLayout()
        
        # 查询条件
        filter_layout = QHBoxLayout()
        
        # 左侧日期选择部分
        date_layout = QHBoxLayout()
        
        # 日期选择
        self.start_date = QDateEdit()
        self.start_date.setDate(QDate.currentDate().addDays(-30))
        self.start_date.setCalendarPopup(True)
        self.start_date.setDisplayFormat("yyyy-MM-dd")
        
        self.end_date = QDateEdit()
        self.end_date.setDate(QDate.currentDate())
        self.end_date.setCalendarPopup(True)
        self.end_date.setDisplayFormat("yyyy-MM-dd")
        
        date_layout.addWidget(QLabel('开始日期:'))
        date_layout.addWidget(self.start_date)
        date_layout.addSpacing(10)
        date_layout.addWidget(QLabel('结束日期:'))
        date_layout.addWidget(self.end_date)
        
        # 右侧按钮部分
        button_layout = QHBoxLayout()
        button_layout.addSpacing(20)
        
        # 查询和导出按钮
        search_btn = QPushButton('查询')
        search_btn.setFixedWidth(80)
        search_btn.clicked.connect(self.search_logs)
        
        export_btn = QPushButton('导出')
        export_btn.setFixedWidth(80)
        export_btn.clicked.connect(self.export_logs)
        
        button_layout.addWidget(search_btn)
        button_layout.addSpacing(10)
        button_layout.addWidget(export_btn)
        
        # 将各部分添加到主工具栏布局
        filter_layout.addLayout(date_layout)
        filter_layout.addStretch()  # 添加弹性空间
        filter_layout.addLayout(button_layout)
        
        # 设置工具栏样式
        toolbar_widget = QWidget()
        toolbar_widget.setLayout(filter_layout)
        toolbar_widget.setStyleSheet("""
            QWidget {
                background-color: #f8f9fa;
                border-bottom: 1px solid #dee2e6;
                padding: 10px;
            }
            QLabel {
                font-size: 14px;
            }
            QPushButton {
                background-color: #007bff;
                color: white;
                border: none;
                padding: 5px 15px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            QDateEdit {
                padding: 5px;
                border: 1px solid #ced4da;
                border-radius: 3px;
                min-width: 120px;
            }
        """)
        
        # 日志表格
        self.log_table = QTableWidget()
        self.log_table.setColumnCount(7)
        self.log_table.setHorizontalHeaderLabels([
            '发送时间', '发件人', '收件人', '收件人姓名', 
            '邮件主题', '发送状态', '错误信息'
        ])
        
        # 设置表格列宽
        header = self.log_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # 发送时间
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)  # 发件人
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)  # 收件人
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # 收件人姓名
        header.setSectionResizeMode(4, QHeaderView.Stretch)           # 邮件主题
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)  # 发送状态
        header.setSectionResizeMode(6, QHeaderView.Stretch)           # 错误信息
        
        # 设置表格样式
        self.log_table.setStyleSheet("""
            QTableWidget {
                background-color: white;
                alternate-background-color: #fafafa;
            }
            QTableWidget::item {
                padding: 5px;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                padding: 5px;
                text-align: center;
            }
        """)
        
        # 设置表格属性
        self.log_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.log_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.log_table.setAlternatingRowColors(True)
        self.log_table.verticalHeader().setVisible(False)
        
        # 添加到主布局
        layout.addWidget(toolbar_widget)
        layout.addWidget(self.log_table)
        self.setLayout(layout)
        
        # 初始加载数据
        self.search_logs()
        
    def search_logs(self):
        """查询日志"""
        try:
            start_date = self.start_date.date().toString('yyyy-MM-dd')
            end_date = self.end_date.date().toString('yyyy-MM-dd')
            
            logs = self.db.get_send_logs(start_date=start_date, end_date=end_date)
            self.log_table.setRowCount(0)
            
            for log in logs:
                row = self.log_table.rowCount()
                self.log_table.insertRow(row)
                
                # 创建表格项并设置居中对齐
                items = [
                    QTableWidgetItem(str(log[8])),  # send_time
                    QTableWidgetItem(str(log[2])),  # sender_email
                    QTableWidgetItem(str(log[3])),  # recipient_email
                    QTableWidgetItem(str(log[4])),  # recipient_name
                    QTableWidgetItem(str(log[5])),  # subject
                    QTableWidgetItem(str(log[6])),  # status
                    QTableWidgetItem(str(log[7] or ''))  # error_message
                ]
                
                # 设置状态列的颜色
                status_color = QColor('#28a745') if log[6] == '成功' else QColor('#dc3545')
                items[5].setForeground(QBrush(status_color))
                
                # 设置单元格对齐方式
                for col, item in enumerate(items):
                    if col in [4, 6]:  # 邮件主题和错误信息左对齐
                        item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                    else:  # 其他列居中对齐
                        item.setTextAlignment(Qt.AlignCenter)
                    self.log_table.setItem(row, col, item)
            
            # 更新窗口标题显示记录数
            self.setWindowTitle(f'发送历史记录 - 共 {len(logs)} 条记录')
            
        except Exception as e:
            MessageBox.show('错误', f'查询日志失败: {str(e)}', 'error', parent=self)
    
    def export_logs(self):
        """导出日志到Excel"""
        try:
            # 获取表格数据
            rows = self.log_table.rowCount()
            cols = self.log_table.columnCount()
            data = []
            
            for row in range(rows):
                row_data = []
                for col in range(cols):
                    item = self.log_table.item(row, col)
                    row_data.append(item.text() if item else '')
                data.append(row_data)
                
            # 创建DataFrame
            df = pd.DataFrame(data, columns=[
                '发送时间', '发件人', '收件人', '收件人姓名',
                '邮件主题', '发送状态', '错误信息'
            ])
            
            # 导出到Excel
            file_name = f'发送历史记录_{QDate.currentDate().toString("yyyy-MM-dd")}.xlsx'
            df.to_excel(file_name, index=False)
            
            MessageBox.show('成功', f'日志已导出到: {file_name}', 'info', parent=self)
        except Exception as e:
            MessageBox.show('错误', f'导出失败: {str(e)}', 'error', parent=self) 