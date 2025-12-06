"""Unified logging functionality using Python's logging module."""

import logging
import os
from logging.handlers import RotatingFileHandler
from ..config.paths import get_log_path

LOG_MAX_BYTES = 3 * 1024 * 1024  # 3 MB per file
LOG_BACKUP_COUNT = 3  # keep 3 backup files

_logger: logging.Logger | None = None


def _get_logger() -> logging.Logger:
    """Get or create the singleton logger instance."""
    global _logger
    if _logger is not None:
        return _logger

    _logger = logging.getLogger("pastemd")
    _logger.setLevel(logging.DEBUG)

    # Avoid adding duplicate handlers
    if _logger.handlers:
        return _logger

    try:
        log_path = get_log_path()
        # Ensure log directory exists
        os.makedirs(os.path.dirname(log_path), exist_ok=True)

        # RotatingFileHandler: auto-rotate when file exceeds max size
        handler = RotatingFileHandler(
            log_path,
            maxBytes=LOG_MAX_BYTES,
            backupCount=LOG_BACKUP_COUNT,
            encoding="utf-8",
        )
        handler.setLevel(logging.DEBUG)

        # Format: [2025-12-07 10:30:00] message
        formatter = logging.Formatter(
            "[%(asctime)s] %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
        handler.setFormatter(formatter)
        _logger.addHandler(handler)
    except Exception:
        # If file handler fails, use NullHandler to prevent errors
        _logger.addHandler(logging.NullHandler())

    return _logger


def log(message: str) -> None:
    """记录日志到文件"""
    try:
        _get_logger().info(message)
    except Exception:
        # 记录日志失败时静默处理，避免递归错误
        pass
