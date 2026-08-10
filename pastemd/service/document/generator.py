"""Document generator - centralized DOCX generation and conversion."""

from typing import Optional, List

from ...integrations.pandoc import PandocIntegration
from ...utils.docx_processor import DocxProcessor
from ...utils.logging import log
from ...core.state import app_state
from ...core.errors import PandocError
from ...config.defaults import DEFAULT_CONFIG
from ...config.loader import ConfigLoader


_DEFAULT_PANDOC_REQUEST_HEADERS: List[str] = [
    "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
]


def _get_pandoc_request_headers(config: dict) -> List[str]:
    if "pandoc_request_headers" not in config:
        return _DEFAULT_PANDOC_REQUEST_HEADERS

    headers = config.get("pandoc_request_headers")
    if headers is None:
        return []
    if isinstance(headers, str):
        return [headers]
    if isinstance(headers, list):
        return [h.strip() for h in headers if isinstance(h, str) and h.strip()]
    return []


def _mask_pandoc_request_headers(headers: List[str]) -> List[str]:
    masked: List[str] = []
    sensitive_names = {
        "authorization",
        "proxy-authorization",
        "cookie",
        "set-cookie",
        "x-api-key",
        "x-auth-token",
    }

    for raw in headers:
        if not isinstance(raw, str):
            continue
        raw = raw.strip()
        if not raw:
            continue

        name, sep, value = raw.partition(":")
        if not sep:
            masked.append(raw[:300] + "...(truncated)" if len(raw) > 300 else raw)
            continue

        header_name = name.strip()
        header_value = value.strip()
        if header_name.lower() in sensitive_names:
            masked.append(f"{header_name}: <redacted>")
            continue

        if len(header_value) > 300:
            header_value = header_value[:300] + "...(truncated)"
        masked.append(f"{header_name}: {header_value}")

    return masked


def _normalize_filters(value) -> List[str]:
    """Normalize filter config values to a flat list of enabled paths.

    Supports both legacy format (list of strings) and new format
    (list of dicts with ``path`` and ``enabled`` keys).  Only paths
    whose ``enabled`` field is truthy (defaults to True) are returned.
    """
    if value is None:
        return []
    if isinstance(value, str):
        return [value.strip()] if value.strip() else []
    if isinstance(value, (list, tuple)):
        result: List[str] = []
        for item in value:
            if isinstance(item, str):
                # Legacy format: plain path string
                if item.strip():
                    result.append(item.strip())
            elif isinstance(item, dict):
                # New format: {"path": "...", "enabled": true/false}
                path = item.get("path", "")
                enabled = item.get("enabled", True)
                if isinstance(path, str) and path.strip() and enabled:
                    result.append(path.strip())
        return result
    return []


def _get_pandoc_filters(config: dict, key: str) -> List[str]:
    global_filters = _normalize_filters(config.get("pandoc_filters"))

    per_filters: List[str] = []
    by_conversion = config.get("pandoc_filters_by_conversion")
    if isinstance(by_conversion, dict):
        per_filters.extend(_normalize_filters(by_conversion.get(key)))

    per_filters.extend(_normalize_filters(config.get(f"pandoc_filters_{key}")))

    combined = []
    seen = set()
    for item in global_filters + per_filters:
        if item not in seen:
            combined.append(item)
            seen.add(item)
    return combined


class DocumentGenerator:
    """
    文档生成服务
    
    负责 Markdown/HTML → DOCX 的转换，管理 Pandoc 初始化与兜底逻辑。
    不做任何通知/UI 操作，只抛出/透传 PandocError 等异常。
    
    Note:
        - 不负责预处理（由 preprocessor 模块负责）
        - 不负责保存文件（由 workflow 或其他模块负责）
        - 只负责纯粹的格式转换
    """
    
    def __init__(self) -> None:
        self._pandoc_integration: Optional[PandocIntegration] = None
    
    def _ensure_pandoc_integration(self) -> None:
        """
        确保 Pandoc 集成已初始化
        
        优先使用 app_state.config["pandoc_path"]，
        失败后回退到 DEFAULT_CONFIG["pandoc_path"]，
        并回写 app_state.config["pandoc_path"] 并保存配置。
        
        Raises:
            PandocError: 如果 Pandoc 初始化失败
        """
        if self._pandoc_integration is not None:
            return
        
        pandoc_path = app_state.config.get("pandoc_path", "pandoc")
        try:
            self._pandoc_integration = PandocIntegration(pandoc_path)
        except PandocError as e:
            log(f"Failed to initialize PandocIntegration: {e}")
            # 回退到默认路径
            try:
                default_path = DEFAULT_CONFIG.get("pandoc_path", "pandoc")
                self._pandoc_integration = PandocIntegration(default_path)
                # 回写配置并保存
                app_state.config["pandoc_path"] = default_path
                config_loader = ConfigLoader()
                config_loader.save(config=app_state.config)
                log(f"Fallback to default pandoc_path: {default_path}")
            except PandocError as e2:
                log(f"Retry to initialize PandocIntegration failed: {e2}")
                self._pandoc_integration = None
                raise PandocError(f"Pandoc initialization failed: {e2}")
    
    def convert_markdown_to_docx_bytes(self, md_text: str, config: dict) -> bytes:
        """
        将 Markdown 文本转换为 DOCX 字节流
        
        Args:
            md_text: 预处理后的 Markdown 文本
            config: 配置字典
            
        Returns:
            DOCX 文件的字节流
            
        Raises:
            PandocError: 转换失败时
            
        Note:
            调用方应该先使用 MarkdownPreprocessor 处理 md_text
        """
        # 1. 转换为 DOCX 字节流
        self._ensure_pandoc_integration()
        request_headers = _get_pandoc_request_headers(config)
        if "pandoc_request_headers" in config and request_headers != _DEFAULT_PANDOC_REQUEST_HEADERS:
            log(
                f"pandoc_request_headers (effective): {_mask_pandoc_request_headers(request_headers)}"
            )
        docx_bytes = self._pandoc_integration.convert_to_docx_bytes(
            md_text=md_text,
            reference_docx=config.get("reference_docx"),
            Keep_original_formula=config.get("Keep_original_formula", False),
            enable_latex_replacements=config.get("enable_latex_replacements", True),
            markdown_hard_line_breaks=config.get("markdown_hard_line_breaks", False),
            custom_filters=_get_pandoc_filters(config, "md_to_docx"),
            request_headers=request_headers,
            cwd=config.get("save_dir"),
        )
        
        # 2. 处理 DOCX 样式
        if config.get("md_disable_first_para_indent", True):
            docx_bytes = DocxProcessor.apply_custom_processing(
                docx_bytes,
                disable_first_para_indent=True,
                target_style="Body Text",
                horizontal_rule_style=config.get("horizontal_rule_style", "default"),
                auto_layout_tables=config.get("docx_auto_table_layout", False),
            )
        elif (
            config.get("horizontal_rule_style") == "paragraph_border"
            or config.get("docx_auto_table_layout", False)
        ):
            docx_bytes = DocxProcessor.apply_custom_processing(
                docx_bytes,
                horizontal_rule_style=config.get("horizontal_rule_style", "default"),
                auto_layout_tables=config.get("docx_auto_table_layout", False),
            )
        
        return docx_bytes
    
    def convert_html_to_docx_bytes(self, html_text: str, config: dict) -> bytes:
        """
        将 HTML 文本转换为 DOCX 字节流
        
        Args:
            html_text: HTML 文本
            config: 配置字典
            
        Returns:
            DOCX 文件的字节流
            
        Raises:
            PandocError: 转换失败时
        """
        # 1. 转换为 DOCX 字节流
        self._ensure_pandoc_integration()
        request_headers = _get_pandoc_request_headers(config)
        if "pandoc_request_headers" in config and request_headers != _DEFAULT_PANDOC_REQUEST_HEADERS:
            log(
                f"pandoc_request_headers (effective): {_mask_pandoc_request_headers(request_headers)}"
            )
        docx_bytes = self._pandoc_integration.convert_html_to_docx_bytes(
            html_text=html_text,
            reference_docx=config.get("reference_docx"),
            Keep_original_formula=config.get("Keep_original_formula", False),
            enable_latex_replacements=config.get("enable_latex_replacements", True),
            custom_filters=_get_pandoc_filters(config, "html_to_docx"),
            custom_filters_html_to_md=_get_pandoc_filters(config, "html_to_md"),
            custom_filters_md_to_docx=_get_pandoc_filters(config, "md_to_docx"),
            request_headers=request_headers,
            cwd=config.get("save_dir"),
        )
        
        # 2. 处理 DOCX 样式
        if config.get("html_disable_first_para_indent", True):
            docx_bytes = DocxProcessor.apply_custom_processing(
                docx_bytes,
                disable_first_para_indent=True,
                target_style="Body Text",
                horizontal_rule_style=config.get("horizontal_rule_style", "default"),
                auto_layout_tables=config.get("docx_auto_table_layout", False),
            )
        elif (
            config.get("horizontal_rule_style") == "paragraph_border"
            or config.get("docx_auto_table_layout", False)
        ):
            docx_bytes = DocxProcessor.apply_custom_processing(
                docx_bytes,
                horizontal_rule_style=config.get("horizontal_rule_style", "default"),
                auto_layout_tables=config.get("docx_auto_table_layout", False),
            )
        
        return docx_bytes

    def convert_html_to_markdown_text(self, html_text: str, config: dict) -> str:
        """
        将 HTML 文本转换为 Markdown 文本（用于富文本粘贴/公式保留链路）。

        Raises:
            PandocError: 转换失败时
        """
        self._ensure_pandoc_integration()
        return self._pandoc_integration.convert_html_to_markdown_text(  # type: ignore[union-attr]
            html_text,
            custom_filters=_get_pandoc_filters(config, "html_to_md"),
        )

    def convert_markdown_to_html_text(self, md_text: str, config: dict) -> str:
        """
        将 Markdown 文本转换为 HTML 文本（用于富文本粘贴）。

        Notes:
            - 通过 Keep_original_formula=True 可把数学节点改成普通文本 `$...$` / `$$...$$`
        """
        self._ensure_pandoc_integration()
        return self._pandoc_integration.convert_markdown_to_html_text(  # type: ignore[union-attr]
            md_text,
            Keep_original_formula=config.get("Keep_original_formula", True),
            enable_latex_replacements=config.get("enable_latex_replacements", True),
            markdown_hard_line_breaks=config.get("markdown_hard_line_breaks", False),
            custom_filters=_get_pandoc_filters(config, "md_to_html"),
            cwd=config.get("save_dir"),
        )

    def convert_markdown_to_rtf_bytes(self, md_text: str, config: dict) -> bytes:
        """
        将 Markdown 文本转换为 RTF 字节流（用于富文本粘贴兜底）。
        """
        self._ensure_pandoc_integration()
        request_headers = _get_pandoc_request_headers(config)
        if "pandoc_request_headers" in config and request_headers != _DEFAULT_PANDOC_REQUEST_HEADERS:
            log(
                f"pandoc_request_headers (effective): {_mask_pandoc_request_headers(request_headers)}"
            )
        return self._pandoc_integration.convert_markdown_to_rtf_bytes(  # type: ignore[union-attr]
            md_text,
            Keep_original_formula=config.get("Keep_original_formula", True),
            enable_latex_replacements=config.get("enable_latex_replacements", True),
            markdown_hard_line_breaks=config.get("markdown_hard_line_breaks", False),
            custom_filters=_get_pandoc_filters(config, "md_to_rtf"),
            request_headers=request_headers,
            cwd=config.get("save_dir"),
        )

    def convert_html_to_latex_text(self, html_text: str, config: dict) -> str:
        """
        将 HTML 文本转换为 LaTeX 文本（用于 Overleaf 粘贴）。
        
        Returns:
            去除文档头部的 LaTeX 内容，可直接粘贴到 Overleaf
        """
        self._ensure_pandoc_integration()
        return self._pandoc_integration.convert_html_to_latex_text(  # type: ignore[union-attr]
            html_text,
            strip_preamble=True,
            custom_filters_html_to_md=_get_pandoc_filters(config, "html_to_md"),
            custom_filters_md_to_latex=_get_pandoc_filters(config, "md_to_latex"),
        )

    def convert_markdown_to_latex_text(self, md_text: str, config: dict) -> str:
        """
        将 Markdown 文本转换为 LaTeX 文本（用于 Overleaf 粘贴）。
        
        Returns:
            去除文档头部的 LaTeX 内容，可直接粘贴到 Overleaf
        """
        self._ensure_pandoc_integration()
        return self._pandoc_integration.convert_markdown_to_latex_text(  # type: ignore[union-attr]
            md_text,
            strip_preamble=True,
            enable_latex_replacements=config.get("enable_latex_replacements", True),
            markdown_hard_line_breaks=config.get("markdown_hard_line_breaks", False),
            custom_filters=_get_pandoc_filters(config, "md_to_latex"),
        )
