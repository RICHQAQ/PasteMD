"""Resource and file path management."""

import os
import sys


def get_base_dir() -> str:
    """获取应用程序基础目录"""
    # 返回项目根目录（pastemd）
    current_file = os.path.abspath(__file__)
    # 从 pastemd/config/paths.py 回到 pastemd/
    return os.path.dirname(os.path.dirname(os.path.dirname(current_file)))


def resource_path(relative_path: str) -> str:
    """获取资源文件路径（支持 PyInstaller）"""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(get_base_dir(), relative_path)


def get_user_data_dir() -> str:
    """获取用户数据目录（跨平台）"""
    if sys.platform == "win32":
        return os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "PasteMD")
    else:
        return os.path.join(os.path.expanduser("~"), ".pastemd")


def ensure_user_data_dir():
    """确保用户数据目录存在"""
    data_dir = get_user_data_dir()
    if not os.path.exists(data_dir):
        os.makedirs(data_dir, exist_ok=True)
    return data_dir


def get_config_path() -> str:
    """获取配置文件路径"""
    data_dir = ensure_user_data_dir()
    return os.path.join(data_dir, "config.json")


def get_log_path() -> str:
    """获取日志文件路径"""
    data_dir = ensure_user_data_dir()
    return os.path.join(data_dir, "pastemd.log")


def get_app_icon_path() -> str:
    """获取应用图标路径 (.ico)"""
    return resource_path(os.path.join("assets", "icons", "logo.ico"))


def get_app_png_path() -> str:
    """获取应用图标路径 (.png)"""
    return resource_path(os.path.join("assets", "icons", "logo.png"))
