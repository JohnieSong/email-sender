from database import Database
import os
import json
import sys
import configparser
from pathlib import Path
from message_box import MessageBox

class Config:
    def __init__(self):
        # 获取用户目录
        user_dir = str(Path.home())
        self.app_dir = os.path.join(user_dir, '.email_sender')
        
        # 创建应用目录
        if not os.path.exists(self.app_dir):
            os.makedirs(self.app_dir)
            
        # 获取程序运行目录（考虑打包后的情况）
        if getattr(sys, 'frozen', False):
            # 打包后的exe路径
            self.exe_dir = os.path.dirname(sys.executable)
        else:
            # 开发环境下的脚本路径
            self.exe_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 设置配置目录
        self.config_dir = os.path.join(self.app_dir, 'config')
        
        # 创建配置目录
        if not os.path.exists(self.config_dir):
            os.makedirs(self.config_dir)
            
        # 设置配置文件路径
        self.settings_file = os.path.join(self.config_dir, 'settings.json')  # 移动到 config 目录
        self.ini_file = os.path.join(self.exe_dir, 'config.ini')  # config.ini 保持在程序目录
        
        # 延迟初始化其他内容
        self._settings = None
        self._db = None
        
        # 创建默认配置文件
        self.create_default_config()
    
    @property
    def settings(self):
        """延迟加载设置"""
        if self._settings is None:
            try:
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    self._settings = json.load(f)
            except:
                self._settings = {}
        return self._settings
    
    def save_settings(self):
        """保存设置到文件"""
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self._settings, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            MessageBox.show_error(f"保存设置失败: {str(e)}")
            return False
    
    @property
    def db(self):
        """延迟加载数据库实例"""
        if self._db is None:
            self._db = Database(self)
        return self._db

    def export_config(self):
        """导出配置到文件"""
        try:
            # 导出模板
            templates = self.db.get_templates()
            with open(os.path.join(self.config_dir, 'templates.json'), 'w', encoding='utf-8') as f:
                json.dump(templates, f, ensure_ascii=False, indent=2)
                
            return True
        except Exception as e:
            MessageBox.show_error(f"导出配置失败: {str(e)}")
            return False

    def import_config(self):
        """从文件导入配置"""
        try:
            # 导入模板
            template_file = os.path.join(self.config_dir, 'templates.json')
            if os.path.exists(template_file):
                with open(template_file, 'r', encoding='utf-8') as f:
                    templates = json.load(f)
                for name, template in templates.items():
                    self.db.add_template(name, template['title'], template['content'])
            
            return True
        except Exception as e:
            MessageBox.show_error(f"导入配置失败: {str(e)}")
            return False

    def get_sender_list(self):
        """获取发件人列表"""
        return self.db.get_sender_list()

    def add_sender(self, email, password, server_type='QQ企业邮箱', nickname=None):
        """添加发件人
        
        Args:
            email: 邮箱地址
            password: 授权码
            server_type: 邮箱类型，默认为QQ企业邮箱
            nickname: 发件人昵称，可选
            
        Returns:
            bool: 是否添加成功
        """
        return self.db.add_sender(email, password, server_type, nickname)

    def remove_sender(self, email):
        """删除发件人"""
        return self.db.remove_sender(email)

    def save_last_sender(self, email):
        """保存最后使用的发件人"""
        try:
            self.settings['last_sender'] = email
            return self.save_settings()
        except Exception as e:
            MessageBox.show_error(f"保存最后使用的发件人失败: {str(e)}")
            return False

    def get_last_sender(self):
        """获取最后使用的发件人"""
        try:
            return self.settings.get('last_sender', '')
        except Exception as e:
            MessageBox.show_error(f"获取最后使用的发件人失败: {str(e)}")
            return ''

    def get_variables(self):
        """从数据库获取所有变量名"""
        return self.db.get_variables()

    def get_test_data(self):
        """从 config.ini 获取测试数据
        
        Returns:
            dict: 包含测试数据的字典，如果配置不存在或读取失败则返回 None
        """
        try:
            # 检查文件是否存在
            if not os.path.exists(self.ini_file):
                MessageBox.show_error("测试配置文件不存在")
                return None
            
            config = configparser.ConfigParser()
            config.read(self.ini_file, encoding='utf-8')
            
            # 检查 Test 部分是否存在
            if not config.has_section('Test'):
                MessageBox.show_error("测试配置不存在")
                return None
            
            # 获取所有测试数据
            test_data = {}
            for key, value in config['Test'].items():
                test_data[key] = value
            
            return test_data
            
        except Exception as e:
            MessageBox.show_error(f"读取测试数据失败: {str(e)}")
            return None

    def save_test_data(self, data):
        """保存测试数据到 config.ini
        
        Args:
            data (dict): 包含测试数据的字典
            
        Returns:
            bool: 保存成功返回 True，失败返回 False
        """
        try:
            config = configparser.ConfigParser()
            
            # 如果文件存在，先读取现有配置
            if os.path.exists(self.ini_file):
                config.read(self.ini_file, encoding='utf-8')
            
            # 确保 Test 部分存在
            if not config.has_section('Test'):
                config.add_section('Test')
            
            # 更新测试数据
            for key, value in data.items():
                config['Test'][key] = str(value)
            
            # 保存到文件
            with open(self.ini_file, 'w', encoding='utf-8') as f:
                config.write(f)
            return True
            
        except Exception as e:
            MessageBox.show_error(f"保存测试数据失败: {str(e)}")
            return False

    def get_smtp_server_list(self):
        """获取SMTP服务器列表
        
        Returns:
            list: 包含所有SMTP服务器信息的列表
        """
        try:
            # 从数据库获取服务器列表
            servers = self.db.get_smtp_servers()
            server_list = []
            for server_type, config in servers.items():
                server_list.append({
                    'server_type': server_type,
                    'config': config
                })
            return server_list
        except Exception as e:
            MessageBox.show_error(f"获取SMTP服务器列表失败: {str(e)}")
            return []

    def save_last_mail_server(self, server_type):
        """保存最后使用的邮件服务器
        
        Args:
            server_type: 服务器类型
        """
        try:
            self.set_config('last_mail_server', server_type)
            self.save_config()
        except Exception as e:
            MessageBox.show_error(f"保存最后使用的邮件服务器失败: {str(e)}")

    def get_last_mail_server(self):
        """获取最后使用的邮件服务器
        
        Returns:
            str: 服务器类型，如果没有则返回 None
        """
        try:
            return self.get_config('last_mail_server')
        except Exception as e:
            MessageBox.show_error(f"获取最后使用的邮件服务器失败: {str(e)}")
            return None

    def create_default_config(self):
        """创建默认的配置文件"""
        try:
            # 创建默认的 settings.json
            if not os.path.exists(self.settings_file):
                default_settings = {
                    'last_sender': '',
                    'last_template': '',
                    'window_size': [1200, 700],
                    'window_position': [100, 100]
                }
                try:
                    with open(self.settings_file, 'w', encoding='utf-8') as f:
                        json.dump(default_settings, f, ensure_ascii=False, indent=2)
                except Exception as e:
                    MessageBox.show_error(f"创建 settings.json 失败: {str(e)}\n路径: {self.settings_file}")
            
            # 创建默认的 config.ini
            if not os.path.exists(self.ini_file):
                try:
                    # 检查是否有写入权限
                    if not os.access(self.exe_dir, os.W_OK):
                        MessageBox.show_error(f"没有写入权限: {self.exe_dir}\n请以管理员身份运行程序")
                        return False
                        
                    config = configparser.ConfigParser()
                    config['Test'] = {
                        'test_sender': '',
                        'test_recipient': '',
                        'test_subject': '',
                        'test_content': ''
                    }
                    with open(self.ini_file, 'w', encoding='utf-8') as f:
                        config.write(f)
                except Exception as e:
                    MessageBox.show_error(f"创建 config.ini 失败: {str(e)}\n路径: {self.ini_file}")
            
            return True
        except Exception as e:
            MessageBox.show_error(f"创建默认配置文件失败: {str(e)}\n应用目录: {self.app_dir}")
            return False

    def get_receiver_list(self):
        """获取收件人列表"""
        return self.db.get_receiver_list()

    def add_receiver(self, email, password, server_type='QQ企业邮箱'):
        """添加收件人"""
        return self.db.add_receiver(email, password, server_type)

    def remove_receiver(self, email):
        """删除收件人"""
        return self.db.remove_receiver(email) 