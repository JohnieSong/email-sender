import os
import sqlite3
import json
from pathlib import Path
from crypto_utils import CryptoUtils
from logger import Logger
import threading

class Database:
    def __init__(self, config=None):
        # 添加线程锁
        self._lock = threading.Lock()
        
        # 获取用户目录
        user_dir = str(Path.home())
        self.app_dir = os.path.join(user_dir, '.email_sender')
        if not os.path.exists(self.app_dir):
            os.makedirs(self.app_dir)
            
        self.db_path = os.path.join(self.app_dir, 'email_sender.db')
        self.crypto = CryptoUtils()
        self.config = config
        
        # 延迟初始化数据库
        self._conn = None
        
        # 初始化日志
        self.logger = Logger()
        
    @property
    def conn(self):
        """延迟创建数据库连接，并设置超时和线程安全"""
        if self._conn is None:
            self._conn = sqlite3.connect(self.db_path, timeout=20)
            # 启用WAL模式以提高并发性能
            self._conn.execute('PRAGMA journal_mode=WAL')
            self.init_database()
        return self._conn

    def init_database(self):
        """初始化数据库"""
        try:
            # 检查数据库文件是否存在
            db_exists = os.path.exists(self.db_path)
            # 创建数据库
            cursor = self.conn.cursor()

            # 创建发件人表
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS senders (
                    email TEXT PRIMARY KEY,
                    password TEXT NOT NULL,
                    server_type TEXT NOT NULL
                )
            ''')

            # 创建SMTP服务器表
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS smtp_servers (
                    server_type TEXT PRIMARY KEY,
                    config TEXT NOT NULL
                )
            ''')

            # 创建模板表
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS templates (
                    name TEXT PRIMARY KEY,
                    title TEXT NOT NULL,
                    content TEXT NOT NULL
                )
            ''')

            # 创建变量表
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS variables (
                    name TEXT PRIMARY KEY,
                    value TEXT
                )
            ''')
            
            # 创建发送日志表
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS send_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    sender_email TEXT NOT NULL,
                    recipient_email TEXT NOT NULL,
                    recipient_name TEXT NOT NULL,
                    subject TEXT NOT NULL,
                    status TEXT NOT NULL,
                    error_message TEXT,
                    send_time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    batch_id TEXT NOT NULL
                )
            ''')
            
            # 创建系统日志表
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS system_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    level TEXT NOT NULL,
                    filename TEXT NOT NULL,
                    line_number INTEGER NOT NULL,
                    message TEXT NOT NULL
                )
            ''')
            
            # 检查SMTP服务器表是否为空
            cursor.execute('SELECT COUNT(*) FROM smtp_servers')
            if cursor.fetchone()[0] == 0:
                # 常见的SMTP服务器配置
                default_servers = [
                    {
                        'server_type': 'QQ邮箱',
                        'config': json.dumps({
                            'smtp_server': 'smtp.qq.com',
                            'smtp_port': 465,
                            'use_ssl': True
                        })
                    },
                    {
                        'server_type': 'QQ企业邮箱',
                        'config': json.dumps({
                            'smtp_server': 'smtp.exmail.qq.com',
                            'smtp_port': 465,
                            'use_ssl': True
                        })
                    },
                    {
                        'server_type': '163邮箱',
                        'config': json.dumps({
                            'smtp_server': 'smtp.163.com',
                            'smtp_port': 465,
                            'use_ssl': True
                        })
                    },
                    {
                        'server_type': '126邮箱',
                        'config': json.dumps({
                            'smtp_server': 'smtp.126.com',
                            'smtp_port': 465,
                            'use_ssl': True
                        })
                    },
                    {
                        'server_type': 'Gmail',
                        'config': json.dumps({
                            'smtp_server': 'smtp.gmail.com',
                            'smtp_port': 587,
                            'use_ssl': False,
                            'use_tls': True
                        })
                    },
                    {
                        'server_type': '阿里企业邮箱',
                        'config': json.dumps({
                            'smtp_server': 'smtp.mxhichina.com',
                            'smtp_port': 465,
                            'use_ssl': True
                        })
                    },
                    {
                        'server_type': '阿里邮箱',
                        'config': json.dumps({
                            'smtp_server': 'smtp.aliyun.com',
                            'smtp_port': 465,
                            'use_ssl': True
                        })
                    },
                    {
                        'server_type': '电信邮箱',
                        'config': json.dumps({
                            'smtp_server': 'smtp.189.cn',
                            'smtp_port': 465,
                            'use_ssl': True
                        })
                    },
                    {
                        'server_type': '搜狐邮箱',
                        'config': json.dumps({
                            'smtp_server': 'smtp.sohu.com',
                            'smtp_port': 465,
                            'use_ssl': True
                        })
                    },
                    {
                        'server_type': '新浪邮箱',
                        'config': json.dumps({
                            'smtp_server': 'smtp.sina.com',
                            'smtp_port': 465,
                            'use_ssl': True
                        })
                    },
                    {
                        'server_type': '移动邮箱',
                        'config': json.dumps({
                            'smtp_server': 'smtp.139.com',
                            'smtp_port': 465,
                            'use_ssl': True
                        })
                    },
                    {
                        'server_type': '苹果邮箱',
                        'config': json.dumps({
                            'smtp_server': 'smtp.mail.me.com',
                            'smtp_port': 587,
                            'use_ssl': False,
                            'use_tls': True
                        })
                    },
                    {
                        'server_type': 'Outlook',
                        'config': json.dumps({
                            'smtp_server': 'smtp.office365.com',
                            'smtp_port': 587,
                            'use_ssl': False,
                            'use_tls': True
                        })
                    },
                    {
                        'server_type': 'AOL邮箱',
                        'config': json.dumps({
                            'smtp_server': 'smtp.aol.com',
                            'smtp_port': 587,
                            'use_ssl': False,
                            'use_tls': True
                        })
                    },
                    {
                        'server_type': 'Yandex邮箱',
                        'config': json.dumps({
                            'smtp_server': 'smtp.yandex.com',
                            'smtp_port': 465,
                            'use_ssl': True
                        })
                    }
                ]
                
                # 插入预设服务器
                for server in default_servers:
                    cursor.execute('''
                        INSERT INTO smtp_servers (server_type, config)
                        VALUES (?, ?)
                    ''', (server['server_type'], server['config']))
                
                self.conn.commit()
                self.logger.info("已初始化默认SMTP服务器配置")

            return True
        except Exception as e:
            self.logger.error(f"初始化数据库失败: {str(e)}")
            return False

    def get_sender_list(self):
        """获取发件人列表"""
        cursor = self.conn.cursor()
        cursor.execute('SELECT email, password, server_type FROM senders')
        senders = []
        for row in cursor.fetchall():
            senders.append({
                'email': row[0],
                'password': self.crypto.decrypt(row[1]),
                'server_type': row[2]
            })
        return senders

    def add_sender(self, email, password, server_type='QQ企业邮箱'):
        """添加或更新发件人"""
        cursor = self.conn.cursor()
        encrypted_password = self.crypto.encrypt(password)
        cursor.execute('''
            INSERT OR REPLACE INTO senders (email, password, server_type)
            VALUES (?, ?, ?)
        ''', (email, encrypted_password, server_type))
        self.conn.commit()
        return True

    def remove_sender(self, email):
        """删除发件人"""
        cursor = self.conn.cursor()
        cursor.execute('DELETE FROM senders WHERE email = ?', (email,))
        self.conn.commit()
        return True

    def get_templates(self):
        """获取所有模板"""
        cursor = self.conn.cursor()
        cursor.execute('SELECT name, title, content FROM templates')
        templates = {}
        for row in cursor.fetchall():
            templates[row[0]] = {
                'title': row[1],
                'content': row[2]
            }
        return templates

    def add_template(self, name, title, content):
        """添加或更新模板"""
        cursor = self.conn.cursor()
        cursor.execute('''
            INSERT OR REPLACE INTO templates (name, title, content)
            VALUES (?, ?, ?)
        ''', (name, title, content))
        self.conn.commit()
        return True

    def remove_template(self, name):
        """删除模板"""
        cursor = self.conn.cursor()
        cursor.execute('DELETE FROM templates WHERE name = ?', (name,))
        self.conn.commit()
        return True

    def update_template(self, name, title, content):
        """更新模板
        
        Args:
            name: 模板名称
            title: 模板标题
            content: 模板内容
            
        Returns:
            bool: 更新是否成功
        """
        try:
            cursor = self.conn.cursor()
            cursor.execute('''
                UPDATE templates 
                SET title = ?, content = ?
                WHERE name = ?
            ''', (title, content, name))
            self.conn.commit()
            return True
        except Exception as e:
            print(f"更新模板失败: {str(e)}")
            return False

    def get_variables(self):
        """获取所有变量
        
        Returns:
            list: 包含所有变量信息的列表，每个变量是一个字典，包含 name 和 value
        """
        try:
            cursor = self.conn.cursor()
            cursor.execute('SELECT name, value FROM variables')
            variables = [{'name': row[0], 'value': row[1]} for row in cursor.fetchall()]
            return variables
        except Exception as e:
            print(f"获取变量失败: {str(e)}")
            return []
    
    def add_variable(self, name, value=''):
        """添加变量
        
        Args:
            name: 变量名
            value: 变量值，默认为空字符串
            
        Returns:
            bool: 添加是否成功
        """
        try:
            cursor = self.conn.cursor()
            cursor.execute('''
                INSERT OR REPLACE INTO variables (name, value) 
                VALUES (?, ?)
            ''', (name, value))
            self.conn.commit()
            return True
        except Exception as e:
            print(f"添加变量失败: {str(e)}")
            return False
    
    def update_variable(self, name, value):
        """更新变量"""
        cursor = self.conn.cursor()
        cursor.execute('''
            INSERT OR REPLACE INTO variables (name, value) 
            VALUES (?, ?)
        ''', (name, value))
        self.conn.commit()
        return True
    
    def get_variable(self, name):
        """获取变量"""
        cursor = self.conn.cursor()
        cursor.execute('SELECT value FROM variables WHERE name = ?', (name,))
        result = cursor.fetchone()
        return result[0] if result else None
    
    def remove_variable(self, name):
        """删除变量"""
        cursor = self.conn.cursor()
        cursor.execute('DELETE FROM variables WHERE name = ?', (name,))
        self.conn.commit()
        return True
    
    def get_smtp_servers(self):
        """获取SMTP服务器配置"""
        cursor = self.conn.cursor()
        cursor.execute('SELECT server_type, config FROM smtp_servers')
        smtp_servers = {}
    
        for row in cursor.fetchall():
            server_type = row[0]
            config = json.loads(row[1])
            smtp_servers[server_type] = config
            
        return smtp_servers
    
    def add_smtp_server(self, server_type, config):
        """添加SMTP服务器配置"""
        cursor = self.conn.cursor()
        cursor.execute('INSERT INTO smtp_servers (server_type, config) VALUES (?, ?)',
                      (server_type, json.dumps(config)))
        self.conn.commit()
        return True
    
    def get_smtp_server(self, server_type):
        """获取指定类型的SMTP服务器配置"""
        cursor = self.conn.cursor()
        cursor.execute('SELECT config FROM smtp_servers WHERE server_type = ?', (server_type,))
        result = cursor.fetchone()
        if result:
            return json.loads(result[0])
        return None

    def update_smtp_server(self, server_type, config):
        """更新SMTP服务器配置"""
        cursor = self.conn.cursor()
        cursor.execute('UPDATE smtp_servers SET config = ? WHERE server_type = ?',
                      (json.dumps(config), server_type))
        self.conn.commit()
        return True

    def remove_smtp_server(self, server_type):
        """删除SMTP服务器配置"""
        cursor = self.conn.cursor()
        cursor.execute('DELETE FROM smtp_servers WHERE server_type = ?', (server_type,))
        self.conn.commit()
        return True

    def get_smtp_server_list(self):
        """获取所有SMTP服务器类型列表
        
        Returns:
            list: 服务器类型列表
        """
        try:
            cursor = self.conn.cursor()
            cursor.execute('SELECT server_type FROM smtp_servers ORDER BY server_type')
            return [row[0] for row in cursor.fetchall()]
        except Exception as e:
            self.logger.error(f"获取SMTP服务器列表失败: {str(e)}")
            # 返回默认的服务器类型列表
            return [
                'QQ企业邮箱',
                'QQ邮箱',
                '163邮箱',
                '126邮箱',
                'Gmail',
                '阿里企业邮箱',
                '自定义'
            ]

    def get_smtp_server_by_name(self, server_name):
        """获取指定名称的SMTP服务器配置"""  
        cursor = self.conn.cursor()
        cursor.execute('SELECT server_type, config FROM smtp_servers WHERE server_type = ?', (server_name,))
        result = cursor.fetchone()
        if result:
            return {
                'server_type': result[0],
                'config': json.loads(result[1])
            }
        return None

    def get_smtp_server_by_domain(self, domain):
        """获取指定域名的SMTP服务器配置"""
        cursor = self.conn.cursor()
        cursor.execute('SELECT server_type, config FROM smtp_servers WHERE config LIKE ?', ('%@' + domain,))
        result = cursor.fetchone()
        if result:
            return {
                'server_type': result[0],
                'config': json.loads(result[1])
            }   
        return None

    def sync_variables_from_config(self):
        """从配置文件同步变量到数据库"""
        try:
            # 从配置文件获取变量列表
            variables = self.config.get_variables()
            
            # 清空现有变量表
            cursor = self.conn.cursor()
            cursor.execute('DELETE FROM variables')
            
            # 插入新变量，设置默认值为空字符串
            for var in variables:
                cursor.execute('''
                    INSERT INTO variables (name, value) 
                    VALUES (?, ?)
                ''', (var, ''))
            
            self.conn.commit()
            return True
        except Exception as e:
            print(f"同步变量失败: {str(e)}")
            return False

    def get_template(self, name):
        """获取指定名称的模板
        
        Args:
            name: 模板名称
            
        Returns:
            dict: 包含模板信息的字典，如果模板不存在则返回 None
        """
        try:
            cursor = self.conn.cursor()
            cursor.execute('SELECT name, title, content FROM templates WHERE name = ?', (name,))
            result = cursor.fetchone()
            
            if result:
                return {
                    'name': result[0],
                    'title': result[1],
                    'content': result[2]
                }
            return None
        except Exception as e:
            print(f"获取模板失败: {str(e)}")
            return None
        
    def add_send_log(self, batch_id, sender_email, recipient_email, recipient_name, subject, status, error_message=None):
        """添加发送日志"""
        cursor = self.conn.cursor()
        cursor.execute('''
            INSERT INTO send_logs (batch_id, sender_email, recipient_email, recipient_name, subject, status, error_message)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (batch_id, sender_email, recipient_email, recipient_name, subject, status, error_message))
        self.conn.commit()
        return True

    def get_send_logs(self, start_date=None, end_date=None, batch_id=None):
        """获取发送日志
        
        Args:
            batch_id (str, optional): 批次ID
            start_date (str, optional): 开始日期，格式：yyyy-MM-dd
            end_date (str, optional): 结束日期，格式：yyyy-MM-dd
            
        Returns:
            list: 发送日志列表
        """
        cursor = self.conn.cursor()
        query = '''
            SELECT 
                id,
                batch_id,
                sender_email,
                recipient_email,
                recipient_name,
                subject,
                status,
                error_message,
                send_time
            FROM send_logs 
            WHERE 1=1
        '''
        params = []
        
        if batch_id:
            query += ' AND batch_id = ?'
            params.append(batch_id)
        if start_date:
            query += ' AND date(send_time) >= date(?)'
            params.append(start_date)
        if end_date:
            # 修改结束日期查询，包含当天的所有数据
            query += ' AND date(send_time) <= date(?, "+1 day")'
            params.append(end_date)
        
        query += ' ORDER BY send_time DESC'
        
        cursor.execute(query, params)
        return cursor.fetchall()

    def add_system_log(self, level, filename, line_number, message):
        """添加系统日志"""
        try:
            with self._lock:
                cursor = self.conn.cursor()
                cursor.execute('''
                    INSERT INTO system_logs (level, filename, line_number, message)
                    VALUES (?, ?, ?, ?)
                ''', (level, filename, line_number, message))
                self.conn.commit()
                return True
        except Exception as e:
            print(f"写入系统日志失败: {str(e)}")
            return False

    def get_system_logs(self, start_date=None, end_date=None, level=None):
        """获取系统日志
        
        Args:
            start_date (str): 开始日期，格式：yyyy-MM-dd
            end_date (str): 结束日期，格式：yyyy-MM-dd
            level (str): 日志级别，可选值：DEBUG, INFO, WARNING, ERROR, CRITICAL
            
        Returns:
            list: [(id, timestamp, level, filename, line_number, message), ...]
        """
        try:
            cursor = self.conn.cursor()
            query = '''
                SELECT id, timestamp, level, filename, line_number, message 
                FROM system_logs 
                WHERE 1=1
            '''
            params = []
            
            if start_date:
                query += ' AND date(timestamp) >= date(?)'
                params.append(start_date)
            if end_date:
                query += ' AND date(timestamp) <= date(?)'
                params.append(end_date)
            if level:
                query += ' AND level = ?'
                params.append(level)
            
            query += ' ORDER BY timestamp DESC'
            
            cursor.execute(query, params)
            return cursor.fetchall()
        except Exception as e:
            self.logger.error(f"获取系统日志失败: {str(e)}")
            return []

    def __del__(self):
        """析构函数：确保在对象销毁时关闭数据库连接"""
        if self._conn:
            try:
                self._conn.close()
            except:
                pass

    def get_senders(self):
        """获取所有发件人列表
        
        Returns:
            list: [(email, password, server_type), ...]
        """
        try:
            cursor = self.conn.cursor()
            cursor.execute('SELECT email, password, server_type FROM senders')
            return cursor.fetchall()
        except Exception as e:
            self.logger.error(f"获取发件人列表失败: {str(e)}")
            return []

    def get_sender(self, email):
        """获取指定发件人信息
        
        Args:
            email (str): 发件人邮箱地址
            
        Returns:
            tuple: (email, password, server_type) 或 None
        """
        try:
            cursor = self.conn.cursor()
            cursor.execute('SELECT email, password, server_type FROM senders WHERE email = ?', (email,))
            return cursor.fetchone()
        except Exception as e:
            self.logger.error(f"获取发件人信息失败: {str(e)}")
            return None

    def update_sender(self, email, password, server_type):
        """更新发件人信息
        
        Args:
            email (str): 发件人邮箱地址
            password (str): 授权码
            server_type (str): 服务器类型
            
        Returns:
            bool: 是否成功
        """
        try:
            cursor = self.conn.cursor()
            cursor.execute('''
                UPDATE senders 
                SET password = ?, server_type = ?
                WHERE email = ?
            ''', (password, server_type, email))
            self.conn.commit()
            return True
        except Exception as e:
            self.logger.error(f"更新发件人失败: {str(e)}")
            return False

    def delete_sender(self, email):
        """删除发件人
        
        Args:
            email (str): 发件人邮箱地址
            
        Returns:
            bool: 是否成功
        """
        try:
            cursor = self.conn.cursor()
            cursor.execute('DELETE FROM senders WHERE email = ?', (email,))
            self.conn.commit()
            return True
        except Exception as e:
            self.logger.error(f"删除发件人失败: {str(e)}")
            return False

    def begin_transaction(self):
        """开始事务"""
        self._lock.acquire()
        self.conn.execute('BEGIN TRANSACTION')

    def commit_transaction(self):
        """提交事务"""
        try:
            self.conn.commit()
        finally:
            self._lock.release()

    def rollback_transaction(self):
        """回滚事务"""
        try:
            self.conn.rollback()
        finally:
            self._lock.release()
        

