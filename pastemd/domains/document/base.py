"""Base class for document inserters."""

from abc import ABC, abstractmethod
from typing import Any, List, Union

from ...utils.com import ensure_com


class BaseDocumentInserter(ABC):
    """文档插入器基类（用于 Word/WPS 文字）"""
    
    def __init__(self, prog_id: Union[str, List[str]], app_name: str):
        """
        初始化插入器
        
        Args:
            prog_id: COM ProgID 或 ProgID 列表 (如 "Word.Application" 或 ["kwps.Application", "KWPS.Application"])
            app_name: 应用名称 (如 "Word" 或 "WPS 文字")
        """
        # 统一转为列表处理
        self.prog_ids = [prog_id] if isinstance(prog_id, str) else prog_id
        self.prog_id = self.prog_ids[0]  # 保持向后兼容
        self.app_name = app_name
    
    @ensure_com
    @abstractmethod
    def insert(self, docx_path: str) -> bool:
        """
        将 DOCX 文件插入到应用当前光标位置
        
        Args:
            docx_path: DOCX 文件路径
            
        Returns:
            True 如果插入成功
            
        Raises:
            InsertError: 插入失败时
        """
        pass
    
    @abstractmethod
    def _get_application(self) -> Any:
        """
        获取应用程序实例
        
        Returns:
            应用程序对象
            
        Raises:
            Exception: 无法获取实例时
        """
        pass
    
    @abstractmethod
    def _perform_insertion(self, app: Any, docx_path: str) -> bool:
        """
        执行实际的插入操作
        
        Args:
            app: 应用程序对象
            docx_path: DOCX 文件路径
            
        Returns:
            True 如果插入成功
            
        Raises:
            InsertError: 插入失败时
        """
        pass
