"""WPS document insertion."""

import time
import win32com.client

from .word import BaseWordInserter
from ...utils.logging import log
from ...utils.win32 import cleanup_background_wps_processes
from ...core.constants import WORD_INSERT_RETRY_COUNT, WORD_INSERT_RETRY_DELAY


class WPSInserter(BaseWordInserter):
    """WPS 文档插入器"""
    
    def __init__(self):
        # WPS 可能有多个不同的 ProgID
        super().__init__(
            prog_id=["kwps.Application", "KWPS.Application"],
            app_name="WPS 文字"
        )

    def insert(self, docx_path):
        """
        插入文档，失败时自动清理后台进程并重试一次
        
        Args:
            docx_path: 文档路径
            
        Returns:
            True 如果插入成功
            
        Raises:
            InsertError: 插入失败时
        """
        try:
            # 第一次尝试
            return super().insert(docx_path)
        except Exception:
            # 第一次失败，尝试清理后台进程
            log("尝试清理后台 WPS 进程重试...")
            cleaned_count = cleanup_background_wps_processes()
            
            if cleaned_count > 0:
                log(f"已清理 {cleaned_count} 个后台 WPS 进程，重试插入...")
                # 清理后重试一次，如果还失败就让异常抛出
                return super().insert(docx_path)
            else:
                log("没有找到需要清理的后台进程")
                raise  # 抛出原始异常

    def _get_application(self):
        """获取 WPS 应用程序实例（尝试所有可能的 ProgID）"""
        for prog_id in self.prog_ids:
            try:
                # 尝试连接现有实例
                app = win32com.client.GetActiveObject(prog_id)
                log(f"Successfully connected to WPS via {prog_id}")
                return app
            except Exception:
                try:
                    # 尝试创建新实例
                    app = win32com.client.Dispatch(prog_id)
                    log(f"Successfully created WPS instance via {prog_id}")
                    return app
                except Exception as e:
                    log(f"Cannot get WPS application via {prog_id}: {e}")
                    continue
        
        raise Exception(f"未找到运行中的 {self.app_name}，请先打开")
    
    def _get_selection(self, app):
        """
        获取 WPS 的选择对象（兼容不同版本）
        
        Args:
            app: WPS 应用程序对象
            
        Returns:
            Selection 对象
            
        Raises:
            Exception: 所有方法都失败时
        """
        import pywintypes
        
        # 方法1：直接从 app 获取 Selection（最常见的方式）
        try:
            selection = app.Selection
            if selection is not None:
                log("获取 WPS Selection 成功（通过 app.Selection）")
                return selection
        except (AttributeError, pywintypes.com_error) as e:
            log(f"无法从 app 获取 Selection: {e}")
        
        # 方法2：通过 ActiveDocument.ActiveWindow.Selection
        try:
            selection = app.ActiveDocument.ActiveWindow.Selection
            if selection is not None:
                log("获取 WPS Selection 成功（通过 ActiveDocument.ActiveWindow.Selection）")
                return selection
        except (AttributeError, pywintypes.com_error) as e:
            log(f"无法从 ActiveDocument.ActiveWindow 获取 Selection: {e}")
        
        # 方法3：通过 ActiveWindow.Selection
        try:
            selection = app.ActiveWindow.Selection
            if selection is not None:
                log("获取 WPS Selection 成功（通过 ActiveWindow.Selection）")
                return selection
        except (AttributeError, pywintypes.com_error) as e:
            log(f"无法从 ActiveWindow 获取 Selection: {e}")
        
        # 方法4：通过 Documents(1).ActiveWindow.Selection
        try:
            documents = app.Documents
            if documents and documents.Count > 0:
                selection = documents(1).ActiveWindow.Selection
                if selection is not None:
                    log("获取 WPS Selection 成功（通过 Documents(1).ActiveWindow.Selection）")
                    return selection
        except (AttributeError, pywintypes.com_error) as e:
            log(f"无法从 Documents(1).ActiveWindow 获取 Selection: {e}")
        
        # 所有方法都失败
        log("所有获取 Selection 的方法都失败")
        raise Exception("无法获取 WPS Selection，可能存在后台进程干扰")
