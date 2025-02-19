import warnings
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
import base64

# 忽略 cryptography 的性能警告
warnings.filterwarnings('ignore', message='You are using cryptography on a 32-bit Python')

class CryptoUtils:
    def __init__(self):
        # 使用固定的盐值（在实际应用中应该安全存储）
        self.salt = b'email_sender_salt' # 盐值 
        self.key = self._generate_key() # 生成加密密钥
        self.fernet = Fernet(self.key) # 创建Fernet对象

    def _generate_key(self):
        """生成加密密钥"""
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(),
            length=32,
            salt=self.salt,
            iterations=100000,
        )
        key = base64.urlsafe_b64encode(kdf.derive(b'email_sender_secret_key'))
        return key

    def encrypt(self, text):
        """加密文本"""
        try:
            return self.fernet.encrypt(text.encode()).decode()
        except:
            return text

    def decrypt(self, encrypted_text):
        """解密文本"""
        try:
            return self.fernet.decrypt(encrypted_text.encode()).decode()
        except:
            return encrypted_text 