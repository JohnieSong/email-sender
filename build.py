import PyInstaller.__main__
import os

# 获取当前目录
base_dir = os.path.abspath(os.path.dirname(__file__))

# 定义图标路径
icon_path = os.path.join(base_dir, "imgs", "logo.ico")

# 构建命令列表
command = [
    'mail_sender.py',  # 主程序文件
    '--name=邮件批量发送工具',  # 生成的exe名称
    '--noconsole',  # 不显示控制台
    '--onefile',  # 打包成单个exe文件
    '--clean',  # 清理临时文件
    '--noupx',  # 不使用UPX压缩
    '--version-file=version_info.txt',  # 版本信息文件
    '--noconfirm',  # 覆盖已存在的文件
    '--log-level=INFO',  # 日志级别
    f'--icon={icon_path}',  # 使用ico图标
    '--add-data=imgs;imgs',  # 添加资源文件
]

# 运行 PyInstaller
PyInstaller.__main__.run(command) 