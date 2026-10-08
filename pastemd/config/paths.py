"""Resource and file path management."""

import os
import sys

from ..utils.system_detect import is_macos, is_windows


def get_base_dir() -> str:
    """获取应用程序基础目录"""
    # 返回项目根目录（pastemd）
    current_file = os.path.abspath(__file__)
    # 从 pastemd/config/paths.py 回到 pastemd/
    return os.path.dirname(os.path.dirname(os.path.dirname(current_file)))


def resource_path(relative_path: str) -> str:
    """
    支持：
    - PyInstaller 单文件 / 非单文件
    - Nuitka 单文件 / 非单文件
    - 源码运行
    """
    if hasattr(sys, "_MEIPASS"):
        # PyInstaller（onefile / onedir）
        base_dir = sys._MEIPASS
    elif getattr(sys, "frozen", False):
        # Nuitka（onefile / standalone）
        base_dir = os.path.dirname(sys.executable)
    else:
        # 源码运行
        base_dir = get_base_dir()

    # 注意用 exists 而不是 isfile：locales 这类目录资源同样需要套 pastemd/ 前缀
    if not os.path.exists(os.path.join(base_dir, relative_path)):
        relative_path_pastemd = os.path.join("pastemd", relative_path)
        if os.path.exists(os.path.join(base_dir, relative_path_pastemd)):
            relative_path = relative_path_pastemd

    return os.path.join(base_dir, relative_path)


def get_user_data_dir() -> str:
    """获取用户数据目录（跨平台）"""
    if is_windows():
        return os.path.join(
            os.environ.get("APPDATA", os.path.expanduser("~")), "PasteMD"
        )
    elif is_macos():
        return os.path.join(
            os.path.expanduser("~"), "Library", "Application Support", "PasteMD"
        )
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


def get_log_dir() -> str:
    if is_macos():
        return os.path.join(os.path.expanduser("~"), "Library", "Logs", "PasteMD")
    else:
        return get_user_data_dir()


def get_log_path() -> str:
    log_dir = get_log_dir()
    os.makedirs(log_dir, exist_ok=True)
    return os.path.join(log_dir, "pastemd.log")


def get_app_icon_path() -> str:
    """获取应用图标路径"""
    if is_macos():
        return resource_path(os.path.join("assets", "icons", "logo.icns"))
    elif is_windows():
        return resource_path(os.path.join("assets", "icons", "logo.ico"))
    else:
        return resource_path(os.path.join("assets", "icons", "logo.png"))


def get_app_white_png_path() -> str:
    """获取应用白色图标路径 (.png)"""
    return resource_path(os.path.join("assets", "icons", "logo_white.png"))


def get_app_png_path() -> str:
    """获取应用图标路径 (.png)"""
    return resource_path(os.path.join("assets", "icons", "logo.png"))


def get_tray_icon_path() -> str:
    """获取托盘图标路径（macOS 使用 logoTemplate.png）"""
    if is_macos():
        return resource_path(os.path.join("assets", "icons", "logoTemplate.png"))
    else:
        return get_app_png_path()


def is_first_launch() -> bool:
    """检测是否为首次启动（通过检查配置文件和日志文件是否存在）"""
    config_path = get_config_path()
    log_path = get_log_path()

    # 如果配置文件和日志文件都不存在，则认为是首次启动
    return not os.path.exists(config_path) and not os.path.exists(log_path)
