from typing import List
from .excel import BaseExcelInserter
from ...utils.logging import log
from ...utils.win32 import cleanup_background_wps_processes
from ...core.errors import InsertError


class WPSExcelInserter(BaseExcelInserter):
    """WPS 表格插入器"""
    
    def __init__(self):
        super().__init__(prog_id=["ket.Application", "et.Application"], app_name="WPS 表格")
    
    def insert(self, table_data: List[List[str]], keep_format: bool = True) -> bool:
        """
        将表格数据插入到 WPS 表格当前光标位置
        
        覆盖基类方法以在获取 ActiveCell 失败时清理后台进程
        
        Args:
            table_data: 二维数组表格数据
            keep_format: 是否保留 Markdown 格式（粗体、斜体等）
            
        Returns:
            True 如果插入成功
            
        Raises:
            InsertError: 插入失败时
        """
        try:
            # 第一次尝试
            return super().insert(table_data, keep_format)
        except InsertError:
            log("尝试清理后台 WPS 进程后重试...")
            cleaned_count = cleanup_background_wps_processes()
            
            if cleaned_count > 0:
                log(f"已清理 {cleaned_count} 个后台 WPS 进程，重试插入...")
                # 清理后重试一次，如果还失败就让异常抛出
                return super().insert(table_data, keep_format)
            else:
                log("没有找到需要清理的后台进程")
                raise  # 抛出原始异常
