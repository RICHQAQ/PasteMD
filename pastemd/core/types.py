"""Common type aliases and protocols."""

from typing import Protocol, Any, Dict, Callable, Literal
from abc import abstractmethod


ConfigDict = Dict[str, Any]
InsertTarget = Literal["auto", "word", "wps", "none"]


class Notifier(Protocol):
    """通知器协议"""
    
    @abstractmethod
    def notify(self, title: str, message: str, ok: bool = True) -> None:
        """发送通知"""
        pass


class Inserter(Protocol):
    """文档插入器协议"""
    
    @abstractmethod
    def insert(self, docx_path: str) -> bool:
        """插入文档到目标应用"""
        pass


class ConfigLoader(Protocol):
    """配置加载器协议"""
    
    @abstractmethod
    def load(self) -> ConfigDict:
        """加载配置"""
        pass
    
    @abstractmethod
    def save(self, config: ConfigDict) -> None:
        """保存配置"""
        pass


HotkeyCallback = Callable[[], None]
