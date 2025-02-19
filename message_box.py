from PyQt5.QtWidgets import QMessageBox

class MessageBox:
    @staticmethod
    def show(title: str, message: str, type: str = "warning", buttons=None, parent=None):
        """统一的消息框管理
        
        Args:
            title: 标题
            message: 消息内容
            type: 消息类型 (warning/info/question/critical)
            buttons: 可选的按钮列表，格式为 [(按钮文本, 角色), ...]
            parent: 父窗口，用于设置消息框的父级关系
        """
        msg = QMessageBox(parent)
        msg.setWindowTitle(title)
        msg.setText(message)
        
        if parent:
            msg.setWindowIcon(parent.windowIcon())
        
        # 根据消息类型设置对应的图标
        if type == "warning":
            msg.setIcon(QMessageBox.Warning)      # 警告图标：黄色感叹号
        elif type == "info":
            msg.setIcon(QMessageBox.Information)  # 信息图标：蓝色字母i
        elif type == "question":
            msg.setIcon(QMessageBox.Question)     # 询问图标：蓝色问号
        elif type == "critical":
            msg.setIcon(QMessageBox.Critical)     # 错误图标：红色叉号
        
        # 设置按钮
        if buttons:
            for text, role in buttons:
                msg.addButton(text, role)
        else:
            msg.addButton("确定", QMessageBox.AcceptRole)
        
        return msg.exec_()

    @staticmethod
    def confirm(title, message, parent=None):
        """显示确认对话框
        
        Args:
            title: 标题
            message: 消息内容
            parent: 父窗口
            
        Returns:
            bool: 用户是否点击了确认按钮
        """
        reply = MessageBox.show(
            title,
            message,
            'question',
            [('确定', QMessageBox.AcceptRole), ('取消', QMessageBox.RejectRole)],
            parent=parent
        )
        return reply == 0  # 0 表示点击了第一个按钮（确定） 