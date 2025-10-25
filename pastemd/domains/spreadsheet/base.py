"""Base class for spreadsheet inserters."""

from abc import ABC, abstractmethod
from typing import Any, List, Union


class BaseTableInserter(ABC):
    """表格插入器基类（用于 Excel/WPS 表格）"""
    
    def __init__(self, prog_id: Union[str, List[str]], app_name: str):
        """
        初始化插入器
        
        Args:
            prog_id: COM ProgID 或 ProgID 列表 (如 "Excel.Application" 或 ["ket.Application"])
            app_name: 应用名称 (如 "Excel" 或 "WPS 表格")
        """
        # 统一转为列表处理
        self.prog_ids = [prog_id] if isinstance(prog_id, str) else prog_id
        self.prog_id = self.prog_ids[0]  # 保持向后兼容
        self.app_name = app_name
    
    @abstractmethod
    def insert(self, table_data: List[List[str]], keep_format: bool = True) -> bool:
        """
        将表格数据插入到应用当前光标位置
        
        Args:
            table_data: 二维数组表格数据
            keep_format: 是否保留 Markdown 格式（粗体、斜体等）
            
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
