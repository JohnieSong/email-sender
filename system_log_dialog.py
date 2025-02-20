from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                            QPushButton, QTableWidget, QTableWidgetItem,
                            QDateEdit, QComboBox, QHeaderView, QWidget)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QColor, QBrush
import pandas as pd
from database import Database
from config import Config
from message_box import MessageBox

class SystemLogDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.config = Config()
        self.db = Database(self.config)
        self.setWindowTitle('系统日志')
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
        self.start_date.setDate(QDate.currentDate().addDays(-7))
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
        
        # 中间日志级别选择部分
        level_layout = QHBoxLayout()
        level_layout.addSpacing(20)
        
        # 日志级别选择
        self.level_combo = QComboBox()
        self.level_combo.addItems(['全部', '调试', '信息', '警告', '错误', '严重'])
        self.level_combo.setFixedWidth(100)  # 设置固定宽度
        
        level_layout.addWidget(QLabel('日志级别:'))
        level_layout.addWidget(self.level_combo)
        
        # 右侧按钮部分
        button_layout = QHBoxLayout()
        button_layout.addSpacing(20)
        
        # 查询和导出按钮
        search_btn = QPushButton('查询')
        search_btn.setFixedWidth(80)  # 设置固定宽度
        search_btn.clicked.connect(self.search_logs)
        
        export_btn = QPushButton('导出')
        export_btn.setFixedWidth(80)  # 设置固定宽度
        export_btn.clicked.connect(self.export_logs)
        
        button_layout.addWidget(search_btn)
        button_layout.addSpacing(10)
        button_layout.addWidget(export_btn)
        
        # 将三个部分添加到主工具栏布局
        filter_layout.addLayout(date_layout)
        filter_layout.addLayout(level_layout)
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
            QComboBox {
                padding: 5px;
                border: 1px solid #ced4da;
                border-radius: 3px;
            }
            QDateEdit {
                padding: 5px;
                border: 1px solid #ced4da;
                border-radius: 3px;
                min-width: 120px;
            }
        """)
        
        # 日志表格部分保持不变
        self.log_table = QTableWidget()
        self.log_table.setColumnCount(3)
        self.log_table.setHorizontalHeaderLabels(['时间', '级别', '内容'])
        
        # 设置表格列宽
        header = self.log_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # 时间
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)  # 级别
        header.setSectionResizeMode(2, QHeaderView.Stretch)           # 内容
        
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
        """查询系统日志"""
        try:
            start_date = self.start_date.date().toString('yyyy-MM-dd')
            end_date = self.end_date.date().toString('yyyy-MM-dd')
            level = self.level_combo.currentText()
            
            # 中文日志级别映射到英文
            level_map = {
                '全部': None,
                '调试': 'DEBUG',
                '信息': 'INFO',
                '警告': 'WARNING',
                '错误': 'ERROR',
                '严重': 'CRITICAL'
            }
            level = level_map.get(level)
            
            logs = self.db.get_system_logs(start_date=start_date, end_date=end_date, level=level)
            self.log_table.setRowCount(0)
            
            # 定义日志级别对应的颜色
            level_colors = {
                'DEBUG': '#6c757d',    # 灰色
                'INFO': '#28a745',     # 绿色
                'WARNING': '#ffc107',  # 黄色
                'ERROR': '#dc3545',    # 红色
                'CRITICAL': '#dc3545'  # 红色
            }
            
            # 英文日志级别映射到中文
            level_display = {
                'DEBUG': '调试',
                'INFO': '信息',
                'WARNING': '警告',
                'ERROR': '错误',
                'CRITICAL': '严重'
            }
            
            for log in logs:
                row = self.log_table.rowCount()
                self.log_table.insertRow(row)
                
                # 格式化日志内容
                message = f"[{log[3]}:{log[4]}] - {log[5]}"  # filename:line_number - message
                
                # 获取日志级别的中文显示
                level = str(log[2])
                level_zh = level_display.get(level, level)
                
                # 创建表格项并设置对齐
                items = [
                    QTableWidgetItem(str(log[1])),  # timestamp
                    QTableWidgetItem(level_zh),     # level in Chinese
                    QTableWidgetItem(message)       # formatted message
                ]
                
                # 设置日志级别的颜色
                level = str(log[2])
                if level in level_colors:
                    items[1].setForeground(QBrush(QColor(level_colors[level])))
                
                # 设置单元格对齐方式
                items[0].setTextAlignment(Qt.AlignCenter)
                items[1].setTextAlignment(Qt.AlignCenter)
                items[2].setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                
                # 设置单元格内容
                for col, item in enumerate(items):
                    self.log_table.setItem(row, col, item)
            
            # 更新窗口标题显示记录数
            self.setWindowTitle(f'系统日志 - 共 {len(logs)} 条记录')
            
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
            df = pd.DataFrame(data, columns=['时间', '级别', '内容'])
            
            # 导出到Excel
            file_name = f'系统日志_{QDate.currentDate().toString("yyyy-MM-dd")}.xlsx'
            df.to_excel(file_name, index=False)
            
            MessageBox.show('成功', f'日志已导出到: {file_name}', 'info', parent=self)
        except Exception as e:
            MessageBox.show('错误', f'导出失败: {str(e)}', 'error', parent=self) 