MODERN_STYLE = """
/* 全局样式 */
QWidget {
    font-family: "Microsoft YaHei", "Segoe UI", "Arial";
    font-size: 12px;
    color: #333333;
}

/* 主窗口 */
QMainWindow {
    background-color: #ffffff;
}

/* 分组框样式 */
QWidget#groupBox {
    background-color: #fafafa;
    border: 1px solid #e8e8e8;
    border-radius: 4px;
}

/* 按钮样式 */
QPushButton {
    background-color: #1890ff;
    color: white;
    border: none;
    padding: 6px 12px;
    border-radius: 3px;
    min-height: 32px;
    font-weight: normal;
    opacity: 0.9;
    font-family: "Microsoft YaHei", "Segoe UI";
    font-size: 14px;
}

QPushButton:hover {
    background-color: #40a9ff;
    opacity: 1;
}

QPushButton:pressed {
    background-color: #096dd9;
}

QPushButton:disabled {
    background-color: #bfbfbf;
    color: #ffffff;
}

/* 输入框样式 */
QLineEdit {
    padding: 6px;
    border: 1px solid #e8e8e8;
    border-radius: 3px;
    background-color: white;
    min-height: 32px;
    font-size: 14px;
    font-family: "Microsoft YaHei", "Segoe UI";
}

QLineEdit:focus {
    border: 1px solid #40a9ff;
}

/* 下拉框样式 */
QComboBox {
    padding: 6px;
    border: 1px solid #e8e8e8;
    border-radius: 3px;
    background-color: white;
    min-height: 32px;
    font-size: 14px;
    font-family: "Microsoft YaHei", "Segoe UI";
}

QComboBox:hover {
    border: 1px solid #40a9ff;
}

/* 标签样式 */
QLabel {
    color: #595959;
    font-weight: normal;
    font-size: 14px;
    font-family: "Microsoft YaHei", "Segoe UI";
}

/* 表格样式 */
QTableWidget {
    border: 1px solid #e8e8e8;
    border-radius: 3px;
    background-color: white;
    gridline-color: #f5f5f5;
}

QTableWidget::item {
    padding: 6px;
}

QTableWidget::item:selected {
    background-color: #e6f7ff;
    color: #333333;
}

QHeaderView::section {
    background-color: #fafafa;
    padding: 6px;
    border: none;
    border-right: 1px solid #e8e8e8;
    border-bottom: 1px solid #e8e8e8;
    font-weight: normal;
}

/* 进度条样式 */
QProgressBar {
    border: 1px solid #e8e8e8;
    border-radius: 3px;
    background-color: #f5f5f5;
    text-align: center;
    min-height: 22px;
}

QProgressBar::chunk {
    background-color: #1890ff;
    border-radius: 2px;
}

/* 状态标签样式 */
QLabel#status_label {
    color: #595959;
    padding: 6px;
    background-color: #fafafa;
    border: 1px solid #e8e8e8;
    border-radius: 3px;
}

/* 文本浏览器样式 */
QTextBrowser {
    border: 1px solid #e8e8e8;
    border-radius: 3px;
    background-color: white;
    padding: 4px;
}

/* 列表部件样式 */
QListWidget {
    border: 1px solid #e8e8e8;
    border-radius: 3px;
    background-color: white;
}

QListWidget::item {
    padding: 4px;
}

QListWidget::item:selected {
    background-color: #e6f7ff;
    color: #333333;
}

/* 分割线样式 */
QSplitter::handle {
    background-color: #e8e8e8;
}

QSplitter::handle:horizontal {
    width: 1px;
}

QSplitter::handle:vertical {
    height: 1px;
}

/* 滚动条样式 */
QScrollBar:vertical {
    border: none;
    background-color: #f5f5f5;
    width: 8px;
    margin: 0px;
}

QScrollBar::handle:vertical {
    background-color: #d9d9d9;
    border-radius: 4px;
    min-height: 20px;
}

QScrollBar::handle:vertical:hover {
    background-color: #bfbfbf;
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}

QScrollBar:horizontal {
    border: none;
    background-color: #f5f5f5;
    height: 8px;
    margin: 0px;
}

QScrollBar::handle:horizontal {
    background-color: #d9d9d9;
    border-radius: 4px;
    min-width: 20px;
}

QScrollBar::handle:horizontal:hover {
    background-color: #bfbfbf;
}

QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
    width: 0px;
}

/* 版权标签样式 */
QLabel#copyright_label {
    color: #a6a6a6;
    padding: 6px;
    font-size: 10px;
    font-family: "Segoe UI", "Microsoft YaHei";
}
""" 