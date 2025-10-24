"""Main paste workflow - orchestrates the entire conversion and insertion process."""

import traceback
import io
import os

from ...utils.win32.detector import detect_active_app
from ...utils.clipboard import get_clipboard_text, is_clipboard_empty
from ...utils.latex import convert_latex_delimiters
from ...domains.awakener import AppLauncher
from ...integrations.pandoc import PandocIntegration
from ...domains.document.word import WordInserter
from ...domains.document.wps import WPSInserter
from ...domains.document.cleanup import FileCleanupManager
from ...domains.spreadsheet.parser import parse_markdown_table
from ...domains.spreadsheet.excel import MSExcelInserter
from ...domains.spreadsheet.wps_excel import WPSExcelInserter
from ...domains.notification.manager import NotificationManager
from ...utils.fs import generate_output_path
from ...utils.logging import log
from ...core.state import app_state
from ...core.errors import ClipboardError, PandocError, InsertError


class PasteWorkflow:
    """转换并插入工作流 - 业务流程编排"""
    
    def __init__(self):
        self.word_inserter = WordInserter()
        self.wps_inserter = WPSInserter()
        self.ms_excel_inserter = MSExcelInserter()
        self.wps_excel_inserter = WPSExcelInserter()
        self.cleanup_manager = FileCleanupManager()
        self.notification_manager = NotificationManager()
        self.pandoc_integration = None  # 延迟初始化
    
    def execute(self) -> None:
        """执行完整的转换和插入流程"""
        try:
            # 1. 检查剪贴板
            if is_clipboard_empty():
                self.notification_manager.notify(
                    "MD2DOCX HotPaste",
                    "剪贴板为空，未处理。",
                    ok=False
                )
                return
            
            # 2. 获取剪贴板内容和配置
            md_text = get_clipboard_text()
            config = app_state.config
            
            # 3. 检测当前活动应用
            target = detect_active_app()
            log(f"Detected active target: {target}")
            
            # 4. 根据目标应用选择处理流程
            if target in ("excel", "wps_excel") and config.get("enable_excel", True):
                # Excel/WPS表格流程：直接插入表格数据
                self._handle_excel_flow(md_text, target, config)
            elif target in ("word", "wps"):
                # Word/WPS文字流程：转换为DOCX后插入
                self._handle_word_flow(md_text, target, config)
            else:
                # 未检测到应用，尝试自动打开预生成的文件
                self._handle_no_app_flow(md_text, config)
            
        except ClipboardError as e:
            log(f"Clipboard error: {e}")
            self.notification_manager.notify(
                "MD2DOCX HotPaste",
                "剪贴板读取失败。",
                ok=False
            )
        except PandocError as e:
            log(f"Pandoc error: {e}")
            self.notification_manager.notify(
                "MD2DOCX HotPaste",
                "Markdown 转换失败，请检查格式。",
                ok=False
            )
        except Exception:
            # 记录详细错误
            error_details = io.StringIO()
            traceback.print_exc(file=error_details)
            log(error_details.getvalue())
            
            self.notification_manager.notify(
                "MD2DOCX HotPaste",
                "转换失败，请查看日志。",
                ok=False
            )
    
    def _handle_excel_flow(self, md_text: str, target: str, config: dict) -> None:
        """
        Excel/WPS表格流程：解析Markdown表格并直接插入
        
        Args:
            md_text: Markdown文本
            target: 目标应用 (excel 或 wps_excel)
            config: 配置字典
        """
        # 根据目标选择插入器
        if target == "wps_excel":
            inserter = self.wps_excel_inserter
            app_name = "WPS 表格"
        else:  # excel
            inserter = self.ms_excel_inserter
            app_name = "Excel"
        
        # 解析Markdown表格
        table_data = parse_markdown_table(md_text)
        
        if table_data is None:
            # 不是有效的Markdown表格
            self.notification_manager.notify(
                "MD2Excel HotPaste",
                f"未检测到有效的 Markdown 表格。\n当前应用: {app_name}",
                ok=False
            )
            return
        
        # 尝试插入表格
        log(f"Detected Markdown table with {len(table_data)} rows, inserting to {app_name}")
        try:
            keep_format = config.get("excel_keep_format", True)
            success = inserter.insert(table_data, keep_format=keep_format)
            
            if success:
                self.notification_manager.notify(
                    "MD2Excel HotPaste",
                    f"已插入 {len(table_data)} 行表格到 {app_name}。",
                    ok=True
                )
        except InsertError as e:
            log(f"{app_name} insert failed: {e}")
            self.notification_manager.notify(
                "MD2Excel HotPaste",
                f"插入到 {app_name} 失败。\n{str(e)}",
                ok=False
            )
    
    def _handle_word_flow(self, md_text: str, target: str, config: dict) -> None:
        """
        Word/WPS文字流程：转换Markdown为DOCX并插入
        
        Args:
            md_text: Markdown文本
            target: 目标应用 (word 或 wps)
            config: 配置字典
        """
        # 1. 处理LaTeX公式
        md_text = convert_latex_delimiters(md_text)
        
        # 2. 生成输出路径
        output_path = generate_output_path(
            keep_file=config.get("keep_file", False),
            save_dir=config.get("save_dir", "")
        )
        
        # 3. 转换为DOCX
        self._ensure_pandoc_integration()
        self.pandoc_integration.convert_to_docx(
            md_text=md_text,
            output_path=output_path,
            reference_docx=config.get("reference_docx")
        )
        log(f"Converted Markdown to DOCX: {output_path}")
        
        # 4. 执行插入
        inserted = self._perform_word_insertion(output_path, target)
        
        # 5. 清理文件
        self.cleanup_manager.cleanup_if_needed(
            file_path=output_path,
            keep_file=config.get("keep_file", False),
            insert_success=inserted,
            target=target
        )
        
        # 6. 显示结果通知
        self._show_word_result(target, inserted)
    
    def _ensure_pandoc_integration(self) -> None:
        """确保 Pandoc 集成已初始化"""
        if self.pandoc_integration is None:
            pandoc_path = app_state.config.get("pandoc_path", "pandoc")
            self.pandoc_integration = PandocIntegration(pandoc_path)
    
    def _perform_word_insertion(self, docx_path: str, target: str) -> bool:
        """
        执行Word/WPS文档插入
        
        Args:
            docx_path: DOCX文件路径
            target: 目标应用 (word 或 wps)
            
        Returns:
            True 如果插入成功
        """
        if target == "word":
            try:
                return self.word_inserter.insert(docx_path)
            except InsertError as e:
                log(f"Word insertion failed: {e}")
                return False
        elif target == "wps":
            try:
                return self.wps_inserter.insert(docx_path)
            except InsertError as e:
                log(f"WPS insertion failed: {e}")
                return False
        else:
            log(f"Unknown insert target: {target}")
            return False
    
    def _show_word_result(self, target: str, inserted: bool) -> None:
        """显示Word/WPS流程的结果通知"""
        if inserted:
            app_name = "Word" if target == "word" else "WPS 文字"
            self.notification_manager.notify(
                "MD2DOCX HotPaste",
                f"已插入到 {app_name}。",
                ok=True
            )
        else:
            app_name = "Word" if target == "word" else "WPS 文字"
            self.notification_manager.notify(
                "MD2DOCX HotPaste",
                f"未能插入到 {app_name}，请确认软件已打开且有光标。",
                ok=False
            )
    
    def _handle_no_app_flow(self, md_text: str, config: dict) -> None:
        """
        无应用检测时的处理流程：生成文件并用默认应用打开
        
        Args:
            md_text: Markdown文本
            config: 配置字典
        """
        # 检查是否启用了自动打开功能
        if not config.get("auto_open_on_no_app", True):
            log("auto_open_on_no_app is disabled, skipping")
            self.notification_manager.notify(
                "MD2DOCX HotPaste",
                "未检测到支持的应用。请打开 Word/WPS/Excel 或启用自动打开。",
                ok=False
            )
            return
        
        # 检测内容类型
        is_table = parse_markdown_table(md_text) is not None
        
        if is_table and config.get("enable_excel", True):
            # 是表格，生成 XLSX 并打开
            self._generate_and_open_spreadsheet(md_text, config)
        else:
            # 是文档，生成 DOCX 并打开
            self._generate_and_open_document(md_text, config)
    
    def _generate_and_open_document(self, md_text: str, config: dict) -> None:
        """
        生成 DOCX 文件并用默认应用打开
        
        Args:
            md_text: Markdown文本
            config: 配置字典
        """
        try:
            # 1. 处理LaTeX公式
            md_text = convert_latex_delimiters(md_text)
            
            # 2. 生成输出路径
            output_path = generate_output_path(
                keep_file=True,  # 生成文件并打开时，默认保留文件
                save_dir=config.get("save_dir", ""),
                md_text=md_text
            )
            
            # 3. 转换为DOCX
            self._ensure_pandoc_integration()
            self.pandoc_integration.convert_to_docx(
                md_text=md_text,
                output_path=output_path,
                reference_docx=config.get("reference_docx")
            )
            log(f"Generated DOCX: {output_path}")
            
            # 4. 用默认应用打开
            if AppLauncher.awaken_and_open_document(output_path):
                self.notification_manager.notify(
                    "MD2DOCX HotPaste",
                    f"已生成文档并用默认应用打开。\n路径: {output_path}",
                    ok=True
                )
            else:
                self.notification_manager.notify(
                    "MD2DOCX HotPaste",
                    f"文档已生成，但打开失败。\n路径: {output_path}",
                    ok=False
                )
        except PandocError as e:
            log(f"Pandoc conversion failed: {e}")
            self.notification_manager.notify(
                "MD2DOCX HotPaste",
                "Markdown 转换失败，请检查格式。",
                ok=False
            )
        except Exception as e:
            log(f"Failed to generate document: {e}")
            self.notification_manager.notify(
                "MD2DOCX HotPaste",
                "生成文档失败。",
                ok=False
            )
    
    def _generate_and_open_spreadsheet(self, md_text: str, config: dict) -> None:
        """
        生成 XLSX 文件并用默认应用打开
        
        Args:
            md_text: Markdown文本
            config: 配置字典
        """
        try:
            # 1. 解析表格
            table_data = parse_markdown_table(md_text)
            if table_data is None:
                self.notification_manager.notify(
                    "MD2DOCX HotPaste",
                    "未检测到有效的 Markdown 表格。",
                    ok=False
                )
                return
            
            # 2. 生成输出路径（XLSX）
            save_dir = config.get("save_dir", "")
            save_dir = os.path.expandvars(save_dir)
            os.makedirs(save_dir, exist_ok=True)
            
            output_path = generate_output_path(
                keep_file=True,
                save_dir=save_dir,
                table_data=table_data
            )
            
            # 3. 生成并打开 XLSX
            keep_format = config.get("excel_keep_format", True)
            if AppLauncher.generate_and_open_spreadsheet(table_data, output_path, keep_format):
                self.notification_manager.notify(
                    "MD2DOCX HotPaste",
                    f"已生成表格（{len(table_data)} 行）并用默认应用打开。\n路径: {output_path}",
                    ok=True
                )
            else:
                self.notification_manager.notify(
                    "MD2DOCX HotPaste",
                    f"表格已生成，但打开失败。\n路径: {output_path}",
                    ok=False
                )
        except Exception as e:
            log(f"Failed to generate spreadsheet: {e}")
            self.notification_manager.notify(
                "MD2DOCX HotPaste",
                "生成表格失败。",
                ok=False
            )
