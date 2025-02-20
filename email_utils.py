import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import re
import os
import time
from email.mime.application import MIMEApplication

class EmailServer:
    # 添加附件大小限制常量（20MB）
    MAX_ATTACHMENT_SIZE = 25 * 1024 * 1024  # 20MB in bytes
    # 添加测试邮件发送间隔（60秒）
    TEST_EMAIL_INTERVAL = 60  # 秒
    # 添加类属性来跟踪最后发送时间
    last_test_time = 0  # 初始化为0
    
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

    def send_email(self, to_email, subject, content, attachments=None):
        """发送邮件"""
        if not self.server:
            self.connect()
            
        # 创建一个带附件的邮件实例
        msg = MIMEMultipart()  # 使用默认的 mixed 类型
        
        # 获取发件人昵称
        senders = self.db.get_sender_list()
        sender_info = None
        for sender in senders:
            if sender['email'] == self.email:
                sender_info = sender
                break
            
        # 设置发件人显示格式
        if sender_info and sender_info.get('nickname'):
            from_addr = f"{sender_info['nickname']} <{self.email}>"
        else:
            from_addr = self.email
        
        msg['From'] = from_addr
        msg['To'] = to_email
        msg['Subject'] = Header(subject, 'utf-8')
        msg['Message-ID'] = f"<{int(time.time())}@{self.email.split('@')[1]}>"
        
        # 创建alternative部分来包含纯文本和HTML内容
        alt_part = MIMEMultipart('alternative')
        
        # 添加纯文本和HTML两种格式
        text_part = MIMEText(self._html_to_text(content), 'plain', 'utf-8')
        html_part = MIMEText(content, 'html', 'utf-8')
        
        alt_part.attach(text_part)
        alt_part.attach(html_part)
        
        # 将alternative部分添加到邮件主体
        msg.attach(alt_part)

        # 如果有附件，添加到邮件中
        if attachments:
            total_size = 0
            for file_path in attachments:
                try:
                    # 检查文件大小
                    file_size = os.path.getsize(file_path)
                    total_size += file_size
                    
                    if file_size > self.MAX_ATTACHMENT_SIZE:
                        raise ValueError(f"单个附件大小超过限制(25MB): {os.path.basename(file_path)}")
                        
                    if total_size > self.MAX_ATTACHMENT_SIZE:
                        raise ValueError(f"附件总大小超过限制(25MB)")
                        
                    # 获取文件名和扩展名
                    filename = os.path.basename(file_path)
                    file_ext = os.path.splitext(filename)[1].lower()
                    
                    # 读取文件内容
                    with open(file_path, 'rb') as f:
                        attachment = f.read()
                    
                    # 根据文件扩展名设置正确的 MIME 类型
                    mime_type = self._get_mime_type(file_ext)
                    
                    # 创建附件部分
                    part = MIMEApplication(attachment)
                    
                    # 设置附件头部信息
                    part.add_header('Content-Disposition', 'attachment', 
                                  filename=('utf-8', '', filename))
                    part.add_header('Content-Type', f'application/octet-stream; name="{filename}"')
                    
                    msg.attach(part)
                    
                except OSError as e:
                    raise ValueError(f"读取附件失败: {file_path} - {str(e)}")
                except Exception as e:
                    raise ValueError(f"添加附件失败: {file_path} - {str(e)}")

        try:
            self.server.send_message(msg)
            self.server.noop()  # 确保邮件发送完成
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
    #         self.close()

    def _get_mime_type(self, file_ext):
        """根据文件扩展名获取 MIME 类型"""
        mime_types = {
            '.pdf': 'pdf',
            '.doc': 'msword',
            '.docx': 'vnd.openxmlformats-officedocument.wordprocessingml.document',
            '.xls': 'vnd.ms-excel',
            '.xlsx': 'vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            '.ppt': 'vnd.ms-powerpoint',
            '.pptx': 'vnd.openxmlformats-officedocument.presentationml.presentation',
            '.jpg': 'jpeg',
            '.jpeg': 'jpeg',
            '.png': 'png',
            '.gif': 'gif',
            '.zip': 'zip',
            '.rar': 'x-rar-compressed',
            '.7z': 'x-7z-compressed',
            '.txt': 'plain',
        }
        return mime_types.get(file_ext, 'octet-stream')