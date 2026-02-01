"""Fallback workflow - handles no-app scenarios."""

from ..base import BaseWorkflow
from .output_executor import OutputExecutor
from pastemd.utils.clipboard import (
    get_clipboard_text, get_clipboard_html, is_clipboard_empty,
    read_markdown_files_from_clipboard
)
from pastemd.utils.html_analyzer import is_plain_html_fragment
from pastemd.utils.markdown_utils import merge_markdown_contents
from pastemd.service.spreadsheet.parser import parse_table
from pastemd.utils.fs import generate_output_path
from pastemd.core.errors import ClipboardError, PandocError
from pastemd.i18n import t


class FallbackWorkflow(BaseWorkflow):
    """无应用场景工作流（执行 no_app_action）"""
    
    def __init__(self):
        super().__init__()
        self.output_executor = OutputExecutor(self.notification_manager)
    
    def execute(self) -> None:
        """
        执行无应用场景工作流
        
        根据剪贴板内容类型和配置决定行为：
        - 如果是表格 → 生成 XLSX
        - 否则 → 生成 DOCX
        - 然后执行 no_app_action (open/save/clipboard)
        """
        content_type: str | None = None
        try:
            # 获取配置的无应用动作
            no_app_action = self.config.get("no_app_action", "open")
            self._log(f"No app detected, executing action: {no_app_action}")
            
            # 1. 检测剪贴板内容类型
            content_type = self._detect_content_type()
            
            # 2. 根据内容类型处理
            if content_type == "table":
                self._handle_table(no_app_action)
            else:
                self._handle_document(no_app_action, content_type)
        
        except ClipboardError as e:
            self._log(f"Clipboard error: {e}")
            msg = str(e)
            if "为空" in msg:
                self._notify_error(t("workflow.clipboard.empty"))
            else:
                self._notify_error(t("workflow.clipboard.read_failed"))
        except PandocError as e:
            self._log(f"Pandoc error: {e}")
            if content_type == "html":
                self._notify_error(t("workflow.html.convert_failed_generic"))
            else:
                self._notify_error(t("workflow.markdown.convert_failed"))
        except Exception as e:
            self._log(f"Fallback workflow failed: {e}")
            import traceback
            traceback.print_exc()
            self._notify_error(t("workflow.generic.failure"))
    
    def _detect_content_type(self) -> str:
        """
        检测剪贴板内容类型
        
        Returns:
            "table" | "html" | "markdown"
        """
        if is_clipboard_empty():
            raise ClipboardError("剪贴板为空")
        
        # 检查是否为表格
        markdown_text = get_clipboard_text()
        found, files_data, _ = read_markdown_files_from_clipboard()
        if found:
            markdown_text = merge_markdown_contents(files_data)
        table_data = parse_table(markdown_text)
        if table_data:
            return "table"
        
        # 检查是否为 HTML
        try:
            html = get_clipboard_html(self.config)
            if not is_plain_html_fragment(html):
                return "html"
        except ClipboardError:
            pass
        
        # 默认为 Markdown
        return "markdown"
    
    def _handle_table(self, action: str):
        """处理表格内容"""
        markdown_text = get_clipboard_text()
        found, files_data, _ = read_markdown_files_from_clipboard()
        if found:
            markdown_text = merge_markdown_contents(files_data)
        table_data = parse_table(markdown_text)
        
        # 生成输出路径
        output_path = generate_output_path(
            keep_file=True,
            save_dir=self.config.get("save_dir", ""),
            table_data=table_data,
        )
        
        # 执行输出
        keep_format = self.config.get("excel_keep_format", self.config.get("keep_format", True))
        success = self.output_executor.execute_xlsx(
            action=action,
            table_data=table_data,
            output_path=output_path,
            keep_format=keep_format
        )
        
        if not success:
            self._log(f"XLSX output failed with action: {action}")
    
    def _handle_document(self, action: str, content_type: str):
        """处理文档内容（HTML 或 Markdown）"""
        # 1. 读取内容
        if content_type == "html":
            html = get_clipboard_html(self.config)
            html = self.html_preprocessor.process(html, self.config)
            docx_bytes = self.doc_generator.convert_html_to_docx_bytes(
                html, self.config
            )
            from_html = True
            md_text = ""
        else:
            # Markdown
            content = get_clipboard_text()
            found, files_data, _ = read_markdown_files_from_clipboard()
            if found:
                content = merge_markdown_contents(files_data)
            # 预处理
            content = self.markdown_preprocessor.process(content, self.config)
            docx_bytes = self.doc_generator.convert_markdown_to_docx_bytes(
                content, self.config
            )
            from_html = False
            md_text = content
        
        # 2. 生成输出路径
        output_path = generate_output_path(
            keep_file=True,
            save_dir=self.config.get("save_dir", ""),
            md_text=md_text,
            html_text=html if from_html else "",
        )
        
        # 3. 执行输出
        success = self.output_executor.execute_docx(
            action=action,
            docx_bytes=docx_bytes,
            output_path=output_path,
            from_md_file=False,
            from_html=from_html
        )
        
        if not success:
            self._log(f"DOCX output failed with action: {action}")
    
    def _read_markdown_content(self) -> str:
        """
        读取 Markdown 内容
        
        Returns:
            Markdown 文本
        """
        # 优先读取文本
        if not is_clipboard_empty():
            return get_clipboard_text()
        
        # 尝试 MD 文件
        found, files_data, _ = read_markdown_files_from_clipboard()
        if found:
            return merge_markdown_contents(files_data)
        
        raise ClipboardError("剪贴板为空或无有效内容")
