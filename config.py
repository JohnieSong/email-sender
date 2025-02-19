from database import Database
import os
import json
import sys
import configparser

class Config:
    def __init__(self):
        # 只初始化必要的路径
        if hasattr(sys, '_MEIPASS'):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.abspath(".")
            
        self.config_dir = os.path.join(base_dir, 'config')
        os.makedirs(self.config_dir, exist_ok=True)
        
        self.settings_file = os.path.join(self.config_dir, 'settings.json')
        self.ini_file = 'config.ini'
        
        # 延迟初始化其他内容
        self._settings = None
        self._db = None
        
    @property
    def db(self):
        """延迟加载数据库实例"""
        if self._db is None:
            self._db = Database(self)
        return self._db

    def _init_settings_file(self):
        """初始化设置文件"""
        if not os.path.exists(self.settings_file):
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump({
                    "last_sender": "",
                    "last_template": "",
                    "window_size": [1400, 900],
                    "splitter_ratio": [25, 75]
                }, f, ensure_ascii=False, indent=2)

    def _init_config_ini(self):
        """初始化配置文件"""
        if not os.path.exists(self.ini_file):
            config = configparser.ConfigParser()
            config['TestData'] = {
                '姓名': '测试用户',
                '考试科目': '测试科目',
                '身份证号': '123456789012345678',
                '考试时间': '2024-01-01 09:00',
                '年龄': '18'  # 添加年龄默认值
            }
            with open(self.ini_file, 'w', encoding='utf-8') as f:
                config.write(f)

    def export_config(self):
        """导出配置到文件"""
        try:
            # 导出模板
            templates = self.db.get_templates()
            with open(os.path.join(self.config_dir, 'templates.json'), 'w', encoding='utf-8') as f:
                json.dump(templates, f, ensure_ascii=False, indent=2)
            
            # 导出发件人（不含密码）
            senders = self.db.get_sender_list()
            safe_senders = [{
                'email': s['email'],
                'server_type': s['server_type']
            } for s in senders]
            with open(os.path.join(self.config_dir, 'senders.json'), 'w', encoding='utf-8') as f:
                json.dump(safe_senders, f, ensure_ascii=False, indent=2)
                
            return True
        except Exception as e:
            print(f"导出配置失败: {str(e)}")
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
            print(f"导入配置失败: {str(e)}")
            return False

    def get_sender_list(self):
        """获取发件人列表"""
        return self.db.get_sender_list()

    def add_sender(self, email, password, server_type='QQ企业邮箱'):
        """添加发件人"""
        return self.db.add_sender(email, password, server_type)

    def remove_sender(self, email):
        """删除发件人"""
        return self.db.remove_sender(email)

    def save_last_sender(self, email):
        """保存最后使用的发件人"""
        try:
            with open(self.settings_file, 'r', encoding='utf-8') as f:
                settings = json.load(f)
            settings['last_sender'] = email
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"保存最后使用的发件人失败: {str(e)}")
            return False

    def get_last_sender(self):
        """获取最后使用的发件人"""
        try:
            with open(self.settings_file, 'r', encoding='utf-8') as f:
                settings = json.load(f)
            return settings.get('last_sender', '')
        except Exception as e:
            print(f"获取最后使用的发件人失败: {str(e)}")
            return ''

    def get_variables(self):
        """从数据库获取所有变量名"""
        return self.db.get_variables()

    def get_test_data(self):
        """从配置文件获取测试数据"""
        try:
            config = configparser.ConfigParser()
            config.read(self.ini_file, encoding='utf-8')
            
            if 'TestData' in config:
                return dict(config['TestData'])
            
            # 如果没有测试数据，从数据库获取变量并创建默认值
            variables = self.get_variables()
            test_data = {var: f'测试{var}' for var in variables}
            self.save_test_data(test_data)
            return test_data
        except Exception as e:
            print(f"读取测试数据失败: {str(e)}")
            return {}

    def save_test_data(self, test_data):
        """保存测试数据到配置文件"""
        try:
            config = configparser.ConfigParser()
            config.read(self.ini_file, encoding='utf-8')
            
            if 'TestData' not in config:
                config['TestData'] = {}
            
            # 更新测试数据
            config['TestData'].update(test_data)
            
            with open(self.ini_file, 'w', encoding='utf-8') as f:
                config.write(f)
            return True
        except Exception as e:
            print(f"保存测试数据失败: {str(e)}")
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
            print(f"获取SMTP服务器列表失败: {str(e)}")
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
            print(f"保存最后使用的邮件服务器失败: {str(e)}")

    def get_last_mail_server(self):
        """获取最后使用的邮件服务器
        
        Returns:
            str: 服务器类型，如果没有则返回 None
        """
        try:
            return self.get_config('last_mail_server')
        except Exception as e:
            print(f"获取最后使用的邮件服务器失败: {str(e)}")
            return None 