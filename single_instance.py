import time
import win32api
import win32event
import win32pipe
import win32file
import winerror
import os
import tempfile
import threading

class SingleInstance:
    """ 确保程序只运行一个实例 """
    def __init__(self, app_window=None):
        self.app_window = app_window
        self.mutexname = "Global\\email_sender_mutex"
        self.pipename = r'\\.\pipe\email_sender_pipe'
        self.lockfile = os.path.join(tempfile.gettempdir(), 'email_sender.lock')
        self.is_running = False
        self.mutex = None
        self.pipe = None
        
        try:
            # 尝试创建互斥锁
            self.mutex = win32event.CreateMutex(None, 1, self.mutexname)
            self.last_error = win32api.GetLastError()
            
            if self.last_error == winerror.ERROR_ALREADY_EXISTS:
                # 互斥锁已存在，说明已有实例在运行
                self.already = True
                return
                
            # 创建命名管道
            self.pipe = win32pipe.CreateNamedPipe(
                self.pipename,
                win32pipe.PIPE_ACCESS_DUPLEX,
                win32pipe.PIPE_TYPE_MESSAGE | win32pipe.PIPE_READMODE_MESSAGE | win32pipe.PIPE_WAIT,
                1, 65536, 65536, 0, None
            )
            
            self.already = False
            
            # 启动管道监听线程
            if self.app_window:
                self.pipe_thread = threading.Thread(target=self._pipe_server, daemon=True)
                self.pipe_thread.start()
            
        except Exception as e:
            print(f"SingleInstance init error: {str(e)}")
            self.already = True
            self.cleanup()

    def cleanup(self):
        """清理资源"""
        if self.pipe:
            try:
                win32file.CloseHandle(self.pipe)
            except:
                pass
            self.pipe = None
            
        if self.mutex:
            try:
                win32api.CloseHandle(self.mutex)
            except:
                pass
            self.mutex = None

    def __del__(self):
        """析构函数"""
        self.cleanup()
        self.cleanup_lockfile()

    def is_first_instance(self):
        """检查是否是第一个实例"""
        return not self.already

    def set_first_instance(self):
        """设置为第一个实例"""
        try:
            # 写入当前进程ID
            with open(self.lockfile, 'w') as f:
                f.write(str(os.getpid()))
            self.is_running = True
        except Exception as e:
            print(f"设置实例失败: {str(e)}")

    def cleanup_lockfile(self):
        """清理锁文件"""
        if self.is_running:
            try:
                os.remove(self.lockfile)
            except:
                pass

    def already_running(self):
        return self.already

    def find_and_activate_window(self):
        """查找并激活已存在的程序窗口"""
        try:
            # 尝试连接到已存在实例的命名管道
            pipe = win32file.CreateFile(
                self.pipename,
                win32file.GENERIC_READ | win32file.GENERIC_WRITE,
                0, None,
                win32file.OPEN_EXISTING,
                0,
                None
            )
            
            # 发送激活消息
            win32file.WriteFile(pipe, b"activate")
            
            # 关闭管道
            win32file.CloseHandle(pipe)
            return True
            
        except Exception as e:
            print(f"激活窗口失败: {str(e)}")
            return False

    def _pipe_server(self):
        """管道服务器线程，处理来自其他实例的请求"""
        while True:
            try:
                # 等待客户端连接
                win32pipe.ConnectNamedPipe(self.pipe, None)
                
                # 读取客户端消息
                hr, data = win32file.ReadFile(self.pipe, 64)
                if data == b"activate":
                    # 使用信号激活窗口
                    if self.app_window:
                        self.app_window.activate_window_signal.emit()
                
                # 断开连接，准备接受下一个连接
                win32pipe.DisconnectNamedPipe(self.pipe)
                
            except Exception as e:
                print(f"Pipe server error: {str(e)}")
                break 