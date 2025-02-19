import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import re

class EmailServer:
    def __init__(self, email, password, server_type=None):
        from database import Database
        self.db = Database()
        
        self.email = email
        self.password = password
        
        # 自动识别邮箱类型
        if server_type is None:
            server_type = self._detect_server_type(email)
            
        self.server_type = server_type
        self.server = None

    def _detect_server_type(self, email):
        """根据邮箱地址自动识别服务器类型"""
        # 从数据库获取服务器配置
        smtp_servers = self.db.get_smtp_servers()
        
        # 获取域名映射关系
        server_map = {}
        for server_type, config in smtp_servers.items():
            if 'domains' in config:
                for domain in config['domains']:
                    server_map[domain] = server_type
        
        domain = email.split('@')[1].lower()
        
        # 检查是否为QQ企业邮箱
        if domain.endswith('exmail.qq.com'):
            return 'QQ企业邮箱'
            
        # 检查是否为Outlook邮箱
        if domain in ['outlook.com', 'hotmail.com']:
            return 'Outlook'
            
        # 从映射中获取服务器类型，如果不存在则返回默认值
        return server_map.get(domain, 'QQ企业邮箱')

    def connect(self):
        """连接到邮件服务器"""
        # 从数据库获取服务器配置
        smtp_servers = self.db.get_smtp_servers()
        
        if self.server_type not in smtp_servers:
            raise ValueError(f'不支持的邮箱类型: {self.server_type}')
            
        server_config = smtp_servers[self.server_type]
        host = server_config['smtp_server']
        port = server_config['smtp_port']
        use_ssl = server_config['use_ssl']
        
        if not use_ssl:
            self.server = smtplib.SMTP(host, port)
            self.server.starttls()  # 使用TLS
        else:
            self.server = smtplib.SMTP_SSL(host, port)
            
        try:
            self.server.login(self.email, self.password)
        except smtplib.SMTPAuthenticationError:
            raise ValueError('邮箱或授权码错误')
        except Exception as e:
            raise ValueError(f'连接服务器失败: {str(e)}')

    def send_email(self, to_email, subject, content):
        """发送邮件"""
        if not self.server:
            self.connect()
            
        msg = MIMEMultipart('alternative')  # 使用 alternative 类型
        msg['From'] = self.email
        msg['To'] = to_email
        msg['Subject'] = Header(subject, 'utf-8')
        
        # 添加纯文本和HTML两种格式
        text_part = MIMEText(self._html_to_text(content), 'plain', 'utf-8')
        html_part = MIMEText(content, 'html', 'utf-8')
        
        msg.attach(text_part)
        msg.attach(html_part)  # HTML 部分会被优先显示

        try:
            self.server.send_message(msg)
            return True, "发送成功"
        except Exception as e:
            return False, str(e)

    def _html_to_text(self, html):
        """将 HTML 转换为纯文本（简单实现）"""
        # 移除 HTML 标签
        text = re.sub(r'<[^>]+>', '', html)
        # 处理特殊字符
        text = text.replace('&nbsp;', ' ')
        text = text.replace('&lt;', '<')
        text = text.replace('&gt;', '>')
        text = text.replace('&amp;', '&')
        return text

    def close(self):
        """关闭连接"""
        if self.server:
            self.server.quit()