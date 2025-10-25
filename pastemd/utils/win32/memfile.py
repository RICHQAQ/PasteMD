# utils/win32/memfile.py
import os, tempfile, time
import win32file, win32con

from pastemd.core.constants import DEFAULT_DELETE_RETRY, DEFAULT_DELETE_WAIT


class EphemeralFile:
    """
    临时文件：允许 READ|WRITE|DELETE 共享，
    以最大兼容 Word/WPS。退出时我们手动删除。
    """
    def __init__(self, suffix=".docx", dir_=None):
        self.dir = dir_ or tempfile.gettempdir()
        os.makedirs(self.dir, exist_ok=True)
        fd, path = tempfile.mkstemp(suffix=suffix, dir=self.dir)
        os.close(fd)
        self.path = path
        self.handle = None

    def write_bytes(self, data: bytes):
        if isinstance(data, str):
            data = data.encode("utf-8")
        # 更宽共享，去掉 DELETE_ON_CLOSE
        self.handle = win32file.CreateFile(
            self.path,
            win32con.GENERIC_WRITE | win32con.GENERIC_READ,
            win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
            None,
            win32con.CREATE_ALWAYS,
            win32con.FILE_ATTRIBUTE_TEMPORARY,   # ← 保留 TEMPORARY 提示缓存
            None
        )
        win32file.WriteFile(self.handle, data)
        # 不强制 Flush，以免触发真正落盘；
        try:
            win32file.SetFilePointerEx(self.handle, 0, win32con.FILE_BEGIN)
        except Exception:
            pass

    def cleanup(self):
        try:
            if self.handle:
                win32file.CloseHandle(self.handle)
                self.handle = None
        except Exception:
            pass
        # 手动删除（多次重试，兼容杀软/索引器短占用）
        for _ in range(DEFAULT_DELETE_RETRY):
            try:
                if os.path.exists(self.path):
                    os.remove(self.path)
                break
            except Exception:
                time.sleep(DEFAULT_DELETE_WAIT)

    def __enter__(self):
        return self

    def __exit__(self, *a):
        self.cleanup()
