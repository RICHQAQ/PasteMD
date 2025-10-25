"""Word document insertion."""

import time
import win32com.client
from win32com.client import gencache

from .base import BaseDocumentInserter
from ...utils.com import ensure_com
from ...utils.logging import log
from ...core.constants import WORD_INSERT_RETRY_COUNT, WORD_INSERT_RETRY_DELAY
from ...core.errors import InsertError


class BaseWordInserter(BaseDocumentInserter):
    """Word 类文档插入器基类（适用于 Word 和 WPS 文字）"""
    
    @ensure_com
    def insert(self, docx_path: str) -> bool:
        """
        将 DOCX 文件插入到当前光标位置
        
        Args:
            docx_path: DOCX 文件路径
            
        Returns:
            True 如果插入成功
            
        Raises:
            InsertError: 插入失败时
        """
        try:
            app = self._get_application()
            return self._perform_insertion(app, docx_path)
        except Exception as e:
            log(f"{self.app_name} insertion failed: {e}")
            raise InsertError(f"{self.app_name} 插入失败: {e}")
    
    def _perform_insertion(self, app, docx_path: str) -> bool:
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
        # 获取当前选择区域
        selection = self._get_selection(app)
        if selection is None:
            raise InsertError(f"无法访问 {self.app_name} 选择区域")
        
        range_obj = selection.Range
        
        # 重试插入文件
        for attempt in range(WORD_INSERT_RETRY_COUNT):
            try:
                range_obj.InsertFile(docx_path)
                log(f"Successfully inserted into {self.app_name}: {docx_path}")
                return True
            except Exception as e:
                if attempt < WORD_INSERT_RETRY_COUNT - 1:
                    log(f"{self.app_name} insert attempt {attempt + 1} failed, retrying: {e}")
                    time.sleep(WORD_INSERT_RETRY_DELAY)
                else:
                    raise InsertError(f"插入失败（已重试 {WORD_INSERT_RETRY_COUNT} 次）: {e}")
        
        return False
    
    def _get_selection(self, app):
        """
        获取选择对象
        
        Args:
            app: 应用程序对象
            
        Returns:
            Selection 对象
        """
        return getattr(app, "Selection", None)
    
    def _ensure_app_ready(self, app) -> None:
        """
        确保应用程序处于就绪状态
        
        Args:
            app: 应用程序对象
        """
        # 默认实现：无需额外操作
        pass

    def _refresh_app(self) -> object:
        """
        刷新应用程序状态（如果需要）
        
        Returns:
            app: 应用程序对象
        """
        return self._get_application()


class WordInserter(BaseWordInserter):
    """Microsoft Word 文档插入器"""
    
    def __init__(self):
        super().__init__(prog_id="Word.Application", app_name="Word")
    
    def _get_application(self):
        """获取 Word 应用程序实例（尝试所有可能的 ProgID）"""
        # 尝试所有可能的 ProgID
        for prog_id in self.prog_ids:
            try:
                # 尝试连接现有的 Word 实例
                app = win32com.client.GetActiveObject(prog_id)
                log(f"Successfully connected to Word via {prog_id}")
                self._ensure_app_ready(app)
                return app
            except Exception:
                try:
                    # 尝试创建新实例
                    app = gencache.EnsureDispatch(prog_id)
                    log(f"Successfully created Word instance via {prog_id}")
                    self._ensure_app_ready(app)
                    return app
                except Exception as e:
                    log(f"Cannot get Word application via {prog_id}: {e}")
                    continue
        
        raise Exception(f"未找到运行中的 {self.app_name}，请先打开")
    
    def _ensure_app_ready(self, app) -> None:
        """确保 Word 应用程序处于就绪状态"""
        try:
            # 确保 Word 可见
            app.Visible = True
        except Exception:
            pass
        
        # 确保有打开的文档
        documents = getattr(app, "Documents", None)
        if documents is None or documents.Count == 0:
            documents.Add()  # 创建新文档
        
        # 切换到主文档正文（避免停留在页眉/页脚/导航窗格）
        try:
            # 0 = wdSeekMainDocument
            app.ActiveWindow.View.SeekView = 0
        except Exception:
            pass
