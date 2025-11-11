"""Main paste workflow - orchestrates the entire conversion and insertion process."""

import traceback
import io
import os
from typing import Optional

from ...utils.win32.detector import detect_active_app
from ...utils.clipboard import get_clipboard_text, is_clipboard_empty, is_clipboard_html, get_clipboard_html
from ...utils.latex import convert_latex_delimiters
from ...domains.awakener import AppLauncher
from ...integrations.pandoc import PandocIntegration
from ...domains.document.word import WordInserter
from ...domains.document.wps import WPSInserter
from ...domains.spreadsheet.parser import parse_markdown_table
from ...domains.spreadsheet.excel import MSExcelInserter
from ...domains.spreadsheet.wps_excel import WPSExcelInserter
from ...domains.notification.manager import NotificationManager
from ...utils.fs import generate_output_path
from ...utils.logging import log
from ...core.state import app_state
from ...core.errors import ClipboardError, PandocError, InsertError
from ...utils.win32.memfile import EphemeralFile
from ...utils.docx_processor import DocxProcessor
from ...utils.html_analyzer import is_plain_html_fragment


class PasteWorkflow:
    """转换并插入工作流 - 业务流程编排"""
    
    def __init__(self):
        self.word_inserter = WordInserter()
        self.wps_inserter = WPSInserter()
        self.ms_excel_inserter = MSExcelInserter()
        self.wps_excel_inserter = WPSExcelInserter()
        self.notification_manager = NotificationManager()
        self.pandoc_integration = None  # 延迟初始化
    
    def execute(self) -> None:
        """执行完整的转换和插入流程"""
        try:
            # 1. 检查剪贴板
            if is_clipboard_empty():
                self.notification_manager.notify(
                    "PasteMD",
                    "剪贴板为空，未处理。",
                    ok=False
                )
                return
            
            # 2. 获取剪贴板内容和配置
            config = app_state.config
            
            # 2.1 检测是否为 HTML 富文本，并尝试识别其结构
            is_html = is_clipboard_html()
            html_text = None
            should_use_html = False
            if is_html:
                try:
                    html_text = get_clipboard_html()
                    is_plain = is_plain_html_fragment(html_text)
                    log(f"Clipboard contains HTML (plain_fragment={is_plain})")
                    if not is_plain:
                        should_use_html = True
                    else:
                        log("HTML fragment looks like Markdown, fallback to Markdown flow.")
                except ClipboardError as e:
                    log(f"Detected HTML clipboard data but failed to read fragment: {e}")
                    is_html = False
            else:
                log("Clipboard contains HTML: False")
            
            # 3. 检测当前活动应用
            target = detect_active_app()
            log(f"Detected active target: {target}")
            
            # 4. 根据剪贴板内容类型和目标应用选择处理流程
            if should_use_html and target in ("word", "wps"):
                # HTML 富文本流程：直接转换 HTML 为 DOCX
                self._handle_html_to_word_flow(target, config, html_text=html_text)
            else:
                # 原有的 Markdown 流程
                md_text = get_clipboard_text()
                
                if target in ("excel", "wps_excel") and config.get("enable_excel", True):
                    # Excel/WPS表格流程：直接插入表格数据
                    self._handle_excel_flow(md_text, target, config)
                elif target in ("word", "wps"):
                    # Word/WPS文字流程：转换为DOCX后插入
                    self._handle_word_flow(md_text, target, config)
                else:
                    # 未检测到应用，尝试自动打开预生成的文件
                    self._handle_no_app_flow(
                        md_text,
                        config,
                        is_html=should_use_html,
                        html_text=html_text if should_use_html else None,
                    )
            
        except ClipboardError as e:
            log(f"Clipboard error: {e}")
            self.notification_manager.notify(
                "PasteMD",
                "剪贴板读取失败。",
                ok=False
            )
        except PandocError as e:
            log(f"Pandoc error: {e}")
            self.notification_manager.notify(
                "PasteMD",
                "Markdown 转换失败，请检查格式。",
                ok=False
            )
        except Exception:
            # 记录详细错误
            error_details = io.StringIO()
            traceback.print_exc(file=error_details)
            log(error_details.getvalue())
            
            self.notification_manager.notify(
                "PasteMD",
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
                "PasteMD",
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
                    "PasteMD",
                    f"已插入 {len(table_data)} 行表格到 {app_name}。",
                    ok=True
                )
        except InsertError as e:
            log(f"{app_name} insert failed: {e}")
            self.notification_manager.notify(
                "PasteMD",
                f"插入到 {app_name} 失败。\n{str(e)}",
                ok=False
            )
    
    def _handle_html_to_word_flow(self, target: str, config: dict, html_text: Optional[str] = None) -> None:
        """
        HTML 富文本流程：直接转换 HTML 为 DOCX 并插入到 Word/WPS
        
        Args:
            target: 目标应用 (word 或 wps)
            config: 配置字典
            html_text: 预先读取好的 HTML 片段（可选）
        """
        try:
            # 1. 获取并清理 HTML 内容
            if html_text is None:
                html_text = get_clipboard_html()
            log(f"Retrieved HTML from clipboard, length: {len(html_text)}")
            
            # 2. 生成 DOCX 字节流
            self._ensure_pandoc_integration()
            docx_bytes = self.pandoc_integration.convert_html_to_docx_bytes(
                html_text=html_text,
                reference_docx=config.get("reference_docx")
            )

            # 3. 在内存中处理 DOCX 样式
            if config.get("html_disable_first_para_indent", True):
                docx_bytes = DocxProcessor.apply_custom_processing(
                    docx_bytes,
                    disable_first_para_indent=True,
                    target_style="Body Text"
                )

            # 4. 使用临时文件插入
            temp_dir = config.get("temp_dir")  # 可选：支持 RAM 盘目录
            with EphemeralFile(suffix=".docx", dir_=temp_dir) as eph:
                eph.write_bytes(docx_bytes)
                # 插入
                inserted = self._perform_word_insertion(eph.path, target)
            
            # 5. 可选保存文件
            if config.get("keep_file", False):
                try:
                    output_path = generate_output_path(
                        keep_file=True,
                        save_dir=config.get("save_dir", ""),
                        html_text=html_text
                    )
                    with open(output_path, "wb") as f:
                        f.write(docx_bytes)
                    log(f"Saved HTML-converted DOCX to: {output_path}")
                except Exception as e:
                    log(f"Failed to save HTML-converted DOCX file: {e}")
            
            # 6. 显示结果通知
            if inserted:
                app_name = "Word" if target == "word" else "WPS 文字"
                self.notification_manager.notify(
                    "PasteMD",
                    f"已从网页 HTML 插入到 {app_name}。",
                    ok=True
                )
            else:
                app_name = "Word" if target == "word" else "WPS 文字"
                self.notification_manager.notify(
                    "PasteMD",
                    f"未能插入到 {app_name}，请确认软件已打开且有光标。",
                    ok=False
                )
                
        except ClipboardError as e:
            log(f"Failed to get HTML from clipboard: {e}")
            self.notification_manager.notify(
                "PasteMD",
                "读取剪贴板 HTML 内容失败。",
                ok=False
            )
        except PandocError as e:
            log(f"HTML to DOCX conversion failed: {e}")
            self.notification_manager.notify(
                "PasteMD",
                "HTML 转换失败，请检查内容格式。",
                ok=False
            )
        except Exception as e:
            log(f"HTML flow failed: {e}")
            error_details = io.StringIO()
            traceback.print_exc(file=error_details)
            log(error_details.getvalue())
            self.notification_manager.notify(
                "PasteMD",
                "HTML 转换失败，请查看日志。",
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

        # 2. 生成DOCX字节流
        self._ensure_pandoc_integration()
        docx_bytes = self.pandoc_integration.convert_to_docx_bytes(
            md_text=md_text,
            reference_docx=config.get("reference_docx")
        )

        # 3. 在内存中处理 DOCX 样式
        if config.get("md_disable_first_para_indent", True):
            docx_bytes = DocxProcessor.apply_custom_processing(
                docx_bytes,
                disable_first_para_indent=True,
                target_style="Body Text"
            )

        # 4. 使用临时文件插入
        temp_dir = config.get("temp_dir")  # 可选：支持 RAM 盘目录
        with EphemeralFile(suffix=".docx", dir_=temp_dir) as eph:
            eph.write_bytes(docx_bytes)
            # 插入
            inserted = self._perform_word_insertion(eph.path, target)

        # 5. 保存文件
        if config.get("keep_file", False):
            # 生成输出路径
            try:
                output_path = generate_output_path(
                    keep_file=config.get("keep_file", False),
                    save_dir=config.get("save_dir", "")
                )
                with open(output_path, "wb") as f:
                    f.write(docx_bytes)
                log(f"Saved DOCX to: {output_path}")
            except Exception as e:
                log(f"Failed to save DOCX file: {e}")
                self.notification_manager.notify(
                    "PasteMD",
                    "保存文档失败。",
                    ok=False
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
                return self.word_inserter.insert(docx_path, app_state.config.get("move_cursor_to_end", True))
            except InsertError as e:
                log(f"Word insertion failed: {e}")
                return False
        elif target == "wps":
            try:
                return self.wps_inserter.insert(docx_path, app_state.config.get("move_cursor_to_end", True))
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
                "PasteMD",
                f"已插入到 {app_name}。",
                ok=True
            )
        else:
            app_name = "Word" if target == "word" else "WPS 文字"
            self.notification_manager.notify(
                "PasteMD",
                f"未能插入到 {app_name}，请确认软件已打开且有光标。",
                ok=False
            )
    
    def _handle_no_app_flow(self, md_text: str, config: dict, is_html: bool = False, html_text: Optional[str] = None) -> None:
        """
        无应用检测时的处理流程：生成文件并用默认应用打开
        支持 Markdown 和 HTML 富文本
        
        Args:
            md_text: Markdown文本
            config: 配置字典
            is_html: 剪贴板是否包含 HTML 富文本
            html_text: 预读取的 HTML 富文本内容
        """
        # 检查是否启用了自动打开功能
        if not config.get("auto_open_on_no_app", True):
            log("auto_open_on_no_app is disabled, skipping")
            self.notification_manager.notify(
                "PasteMD",
                "未检测到支持的应用。请打开 Word/WPS/Excel 或启用自动打开。",
                ok=False
            )
            return
        
        # HTML 富文本优先处理
        if is_html:
            self._generate_and_open_html_document(md_text, config, html_text=html_text)
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

            # 3. 转换为DOCX字节流
            self._ensure_pandoc_integration()
            docx_bytes = self.pandoc_integration.convert_to_docx_bytes(
                md_text=md_text,
                reference_docx=config.get("reference_docx")
            )

            # 4. 在内存中处理 DOCX 样式
            if config.get("md_disable_first_para_indent", True):
                docx_bytes = DocxProcessor.apply_custom_processing(
                    docx_bytes,
                    disable_first_para_indent=True,
                    target_style="Body Text"
                )

            # 5. 写入文件
            with open(output_path, "wb") as f:
                f.write(docx_bytes)
            log(f"Generated DOCX: {output_path}")

            # 6. 用默认应用打开
            if AppLauncher.awaken_and_open_document(output_path):
                self.notification_manager.notify(
                    "PasteMD",
                    f"已生成文档并用默认应用打开。\n路径: {output_path}",
                    ok=True
                )
            else:
                self.notification_manager.notify(
                    "PasteMD",
                    f"文档已生成，但打开失败。\n路径: {output_path}",
                    ok=False
                )
        except PandocError as e:
            log(f"Pandoc conversion failed: {e}")
            self.notification_manager.notify(
                "PasteMD",
                "Markdown 转换失败，请检查格式。",
                ok=False
            )
        except Exception as e:
            log(f"Failed to generate document: {e}")
            self.notification_manager.notify(
                "PasteMD",
                "生成文档失败。",
                ok=False
            )
    
    def _generate_and_open_html_document(self, md_text: str, config: dict, html_text: Optional[str] = None) -> None:
        """生成 DOCX 文件（来源 HTML）并用默认应用打开"""
        try:
            if html_text is None:
                html_text = get_clipboard_html()
            log(f"Retrieved HTML from clipboard for auto-open, length: {len(html_text)}")

            self._ensure_pandoc_integration()
            docx_bytes = self.pandoc_integration.convert_html_to_docx_bytes(
                html_text=html_text,
                reference_docx=config.get("reference_docx")
            )

            if config.get("html_disable_first_para_indent", True):
                docx_bytes = DocxProcessor.apply_custom_processing(
                    docx_bytes,
                    disable_first_para_indent=True,
                    target_style="Body Text"
                )

            output_path = generate_output_path(
                keep_file=True,
                save_dir=config.get("save_dir", ""),
                md_text=md_text,
                html_text=html_text
            )

            with open(output_path, "wb") as f:
                f.write(docx_bytes)
            log(f"Generated DOCX from HTML: {output_path}")

            if AppLauncher.awaken_and_open_document(output_path):
                self.notification_manager.notify(
                    "PasteMD",
                    f"已从 HTML 生成文档并用默认应用打开。\n路径: {output_path}",
                    ok=True
                )
            else:
                self.notification_manager.notify(
                    "PasteMD",
                    f"文档已生成，但打开失败。\n路径: {output_path}",
                    ok=False
                )
        except ClipboardError as e:
            log(f"Failed to get HTML from clipboard: {e}")
            self.notification_manager.notify(
                "PasteMD",
                "读取剪贴板 HTML 内容失败。",
                ok=False
            )
        except PandocError as e:
            log(f"HTML to DOCX conversion failed: {e}")
            self.notification_manager.notify(
                "PasteMD",
                "HTML 转换失败，请检查内容格式。",
                ok=False
            )
        except Exception as e:
            log(f"Failed to generate HTML document: {e}")
            self.notification_manager.notify(
                "PasteMD",
                "生成 HTML 文档失败。",
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
                    "PasteMD",
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
                    "PasteMD",
                    f"已生成表格（{len(table_data)} 行）并用默认应用打开。\n路径: {output_path}",
                    ok=True
                )
            else:
                self.notification_manager.notify(
                    "PasteMD",
                    f"表格已生成，但打开失败。\n路径: {output_path}",
                    ok=False
                )
        except Exception as e:
            log(f"Failed to generate spreadsheet: {e}")
            self.notification_manager.notify(
                "PasteMD",
                "生成表格失败。",
                ok=False
            )
