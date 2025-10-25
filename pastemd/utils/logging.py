"""Unified logging functionality."""

from datetime import datetime
from ..config.paths import get_log_path


def log(message: str) -> None:
    """记录日志到文件"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_line = f"[{timestamp}] {message}\n"
    
    try:
        with open(get_log_path(), "a", encoding="utf-8") as f:
            f.write(log_line)
    except Exception:
        # 记录日志失败时静默处理，避免递归错误
        pass
