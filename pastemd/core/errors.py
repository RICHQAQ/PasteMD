"""Custom exceptions for the application."""


class PasteMDError(Exception):
    """应用程序基础异常"""
    pass


class ConfigError(PasteMDError):
    """配置相关异常"""
    pass


class PandocError(PasteMDError):
    """Pandoc 转换异常"""
    pass


class InsertError(PasteMDError):
    """文档插入异常"""
    pass


class ClipboardError(PasteMDError):
    """剪贴板操作异常"""
    pass
