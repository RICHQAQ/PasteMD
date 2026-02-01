"""Pandoc CLI tool integration."""

import os
import re
import subprocess
from typing import Optional, List

from ..utils.html_formatter import protect_brackets

from ..config.paths import resource_path

from ..core.errors import PandocError
from ..utils.logging import log

LUA_KEEP_ORIGINAL_FORMULA = resource_path("lua/keep-latex-math.lua")
LUA_LATEX_REPLACEMENTS = resource_path("lua/latex-replacements.lua")


def _log_pandoc_stderr_as_warning(stderr: Optional[bytes], *, context: str) -> None:
    if not stderr:
        return

    msg = stderr.decode("utf-8", "ignore").strip()
    if not msg:
        return

    max_len = 4000
    if len(msg) > max_len:
        msg = msg[:max_len] + "...(truncated)"

    log(f"{context} (stderr): {msg}")


def _add_request_headers(cmd: List[str], request_headers: Optional[List[str]]) -> List[str]:
    if not request_headers:
        return cmd
    for header in request_headers:
        if not isinstance(header, str):
            continue
        header = header.strip()
        if not header:
            continue
        cmd += ["--request-header", header]
    return cmd


class PandocIntegration:
    """Pandoc 工具集成"""
    
    def __init__(self, pandoc_path: str = "pandoc"):
        # 测试 Pandoc 可执行文件路径
        cmd = [pandoc_path, "--version"]
        try:
            startupinfo = None
            creationflags = 0
            if os.name == "nt":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                creationflags = subprocess.CREATE_NO_WINDOW
            
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                shell=False,
                startupinfo=startupinfo,
                creationflags=creationflags,
            )
            if result.returncode != 0:
                raise PandocError(f"Pandoc not found or not working: {result.stderr.strip()}")
        except FileNotFoundError:
            raise PandocError(f"Pandoc executable not found: {pandoc_path}")
        except Exception as e:
            raise PandocError(f"Pandoc Error: {e}")
        self.pandoc_path = pandoc_path

    def _build_filter_args(self, custom_filters: Optional[List[str]] = None) -> List[str]:
        """
        构建 Pandoc Filter 参数列表
        
        Args:
            custom_filters: 自定义 Filter 文件路径列表
            
        Returns:
            Filter 参数列表，格式为 ["--lua-filter", "path1", "--filter", "path2", ...]
        """
        filter_args = []
        
        if not custom_filters:
            return filter_args
        
        for filter_path in custom_filters:
            # 展开环境变量
            expanded_path = os.path.expandvars(filter_path)
            
            # 检查文件是否存在
            if not os.path.isabs(expanded_path):
                # 相对路径转换为绝对路径（相对于当前工作目录）
                expanded_path = os.path.abspath(expanded_path)
            
            if not os.path.exists(expanded_path):
                log(f"Warning: Filter file not found, skipping: {expanded_path}")
                continue
            
            # 根据文件扩展名选择参数类型
            if expanded_path.lower().endswith('.lua'):
                filter_args.extend(["--lua-filter", expanded_path])
            else:
                filter_args.extend(["--filter", expanded_path])
        
        return filter_args

    def _convert_html_to_md(
        self,
        html_text: str,
        custom_filters: Optional[List[str]] = None,
    ) -> str:
        """
        使用 Pandoc 将 HTML 转换为 Markdown。
        """
        html_text = protect_brackets(html_text)
        cmd = [
            self.pandoc_path,
            "-f", "html+tex_math_dollars+raw_tex+tex_math_double_backslash+tex_math_single_backslash",
            "-t", "gfm-raw_html+tex_math_dollars",
            "-o", "-",          # 输出到 stdout
            "--wrap", "none",   # 不自动换行，方便你后处理
        ]
        cmd += self._build_filter_args(custom_filters)

        startupinfo = None
        creationflags = 0
        if os.name == "nt":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            creationflags = subprocess.CREATE_NO_WINDOW

        result = subprocess.run(
            cmd,
            input=html_text.encode("utf-8"),  # 显式用 UTF-8 编码
            capture_output=True,
            text=False,                       # 二进制模式
            shell=False,
            startupinfo=startupinfo,
            creationflags=creationflags,
        )
        if result.returncode != 0:
            err = (result.stderr or b"").decode("utf-8", "ignore")
            log(f"Pandoc HTML to MD error: {err}")
            raise PandocError(err or "Pandoc HTML to Markdown conversion failed")

        # stdout 也是 bytes，自行按 UTF-8 解码
        md = result.stdout.decode("utf-8", "ignore")
        md = md.replace('\r\n', '\n').replace('\r', '\n')  # 统一换行符
        md = re.sub(r'```\s*math\s*\n(.*?)\n\s*```', r'$$\n\1\n$$', md, flags=re.DOTALL)
        md = re.sub(r'\$\s*`([^`]+)`\s*\$', r'$\1$', md)
        md = re.sub(r'(```\s*\w+)\s+[^\n]+', r'\1', md)
        # 处理\~~删除线文本\~~
        md = re.sub(r'\\~~(.*?)\\~~', r'~~\1~~', md)

        md = md.replace("{{TASK_CHECKED}}", "[x]").replace("{{TASK_UNCHECKED}}", "[ ]")
        return md

    def convert_html_to_markdown_text(
        self,
        html_text: str,
        *,
        custom_filters: Optional[List[str]] = None,
    ) -> str:
        """
        将 HTML 转换为 Markdown 文本（保留 $...$ 数学语法）。
        """
        return self._convert_html_to_md(html_text, custom_filters)

    def convert_markdown_to_html_text(
        self,
        md_text: str,
        *,
        Keep_original_formula: bool = False,
        enable_latex_replacements: bool = True,
        custom_filters: Optional[List[str]] = None,
        cwd: Optional[str] = None,
    ) -> str:
        """
        将 Markdown 转换为 HTML 文本（用于富文本粘贴）。

        Note:
            - Keep_original_formula=True 时，会用 keep-latex-math.lua 将数学节点改成普通文本 `$...$` / `$$...$$`。
            - 输出为 HTML fragment
        """
        cmd = [
            self.pandoc_path,
            "-f", "markdown+tex_math_dollars+raw_tex+tex_math_double_backslash+tex_math_single_backslash",
            "-t", "html",
            "-o", "-",
            "--wrap", "none",
            "--standalone",
            "--mathml",
        ]
        if enable_latex_replacements:
            cmd += ["--lua-filter", LUA_LATEX_REPLACEMENTS]
        if Keep_original_formula:
            cmd += ["--lua-filter", LUA_KEEP_ORIGINAL_FORMULA]
        cmd += self._build_filter_args(custom_filters)

        # 确保工作目录存在且可写
        if cwd:
            cwd = os.path.expandvars(cwd)
            os.makedirs(cwd, exist_ok=True)

        startupinfo = None
        creationflags = 0
        if os.name == "nt":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            creationflags = subprocess.CREATE_NO_WINDOW

        result = subprocess.run(
            cmd,
            input=md_text.encode("utf-8"),
            capture_output=True,
            text=False,
            shell=False,
            startupinfo=startupinfo,
            creationflags=creationflags,
            cwd=cwd,
        )
        if result.returncode != 0:
            err = (result.stderr or b"").decode("utf-8", "ignore")
            log(f"Pandoc Markdown to HTML error: {err}")
            raise PandocError(err or "Pandoc Markdown to HTML conversion failed")

        return result.stdout.decode("utf-8", "ignore")

    def convert_markdown_to_rtf_bytes(
        self,
        md_text: str,
        *,
        Keep_original_formula: bool = False,
        enable_latex_replacements: bool = True,
        custom_filters: Optional[List[str]] = None,
        request_headers: Optional[List[str]] = None,
        cwd: Optional[str] = None,
    ) -> bytes:
        """
        将 Markdown 转换为 RTF 字节（用于富文本粘贴兜底）。
        """
        cmd = [
            self.pandoc_path,
            "-f", "markdown+tex_math_dollars+raw_tex+tex_math_double_backslash+tex_math_single_backslash",
            "-t", "rtf",
            "-o", "-",
            "--standalone",
        ]
        if enable_latex_replacements:
            cmd += ["--lua-filter", LUA_LATEX_REPLACEMENTS]
        if Keep_original_formula:
            cmd += ["--lua-filter", LUA_KEEP_ORIGINAL_FORMULA]
        cmd += self._build_filter_args(custom_filters)
        cmd = _add_request_headers(cmd, request_headers)

        # 确保工作目录存在且可写
        if cwd:
            cwd = os.path.expandvars(cwd)
            os.makedirs(cwd, exist_ok=True)

        startupinfo = None
        creationflags = 0
        if os.name == "nt":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            creationflags = subprocess.CREATE_NO_WINDOW

        result = subprocess.run(
            cmd,
            input=md_text.encode("utf-8"),
            capture_output=True,
            text=False,
            shell=False,
            startupinfo=startupinfo,
            creationflags=creationflags,
            cwd=cwd,
        )
        if result.returncode != 0:
            err = (result.stderr or b"").decode("utf-8", "ignore")
            log(f"Pandoc Markdown to RTF error: {err}")
            raise PandocError(err or "Pandoc Markdown to RTF conversion failed")

        return result.stdout

    def convert_to_docx_bytes(self, md_text: str, reference_docx: Optional[str] = None, Keep_original_formula: bool = False, enable_latex_replacements: bool = True, custom_filters: Optional[List[str]] = None, request_headers: Optional[List[str]] = None, cwd: Optional[str] = None) -> bytes:
        """
        用 stdin 喂入 Markdown，直接把 DOCX 从 stdout 读到内存（无任何输入文件写盘）
        
        Args:
            md_text: Markdown 文本
            reference_docx: 参考文档模板路径
            Keep_original_formula: 是否保留原始公式
            enable_latex_replacements: 是否启用 LaTeX 替换
            custom_filters: 自定义 Filter 列表
            cwd: Pandoc 进程的工作目录，用于 Filter 创建临时文件（如 mermaid-filter.err）
            
        Returns:
            DOCX 文件的字节流
        """
        cmd = [
            self.pandoc_path,
            "-f", "markdown+tex_math_dollars+raw_tex+tex_math_double_backslash+tex_math_single_backslash",
            "-t", "docx",
            "-o", "-",
            "--highlight-style", "tango",
        ]
        if enable_latex_replacements:
            cmd += ["--lua-filter", LUA_LATEX_REPLACEMENTS]
        if Keep_original_formula:
            cmd += ["--lua-filter", LUA_KEEP_ORIGINAL_FORMULA]
        # 添加自定义 Filter
        cmd += self._build_filter_args(custom_filters)
        if reference_docx:
            cmd += ["--reference-doc", reference_docx]
        cmd = _add_request_headers(cmd, request_headers)

        startupinfo = None
        creationflags = 0
        if os.name == "nt":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            creationflags = subprocess.CREATE_NO_WINDOW

        # 确保工作目录存在且可写
        if cwd:
            cwd = os.path.expandvars(cwd)
            os.makedirs(cwd, exist_ok=True)

        # 关键：input 直接传 UTF-8 字节；text=False 以得到二进制 stdout
        result = subprocess.run(
            cmd,
            input=md_text.encode("utf-8"),
            capture_output=True,
            text=False,
            shell=False,
            startupinfo=startupinfo,
            creationflags=creationflags,
            cwd=cwd,
        )
        if result.returncode != 0:
            # stderr 可能是字节，转成字符串便于日志查看
            err = (result.stderr or b"").decode("utf-8", "ignore")
            log(f"Pandoc error: {err}")
            raise PandocError(err or "Pandoc conversion failed")

        _log_pandoc_stderr_as_warning(result.stderr, context="Pandoc warning (MD->DOCX)")
        return result.stdout

    def convert_html_to_docx_bytes(
        self,
        html_text: str,
        reference_docx: Optional[str] = None,
        Keep_original_formula: bool = False,
        enable_latex_replacements: bool = True,
        custom_filters: Optional[List[str]] = None,
        request_headers: Optional[List[str]] = None,
        cwd: Optional[str] = None,
        custom_filters_html_to_md: Optional[List[str]] = None,
        custom_filters_md_to_docx: Optional[List[str]] = None,
    ) -> bytes:
        """
        用 stdin 喂入 HTML，直接把 DOCX 从 stdout 读到内存（无任何输入文件写盘）
        
        Args:
            html_text: HTML 文本内容
            reference_docx: 可选的参考文档模板路径
            Keep_original_formula: 是否保留原始公式
            enable_latex_replacements: 是否启用 LaTeX 替换
            custom_filters: 自定义 Filter 列表
            cwd: Pandoc 进程的工作目录，某些 Filter 可能会在此目录下创建临时文件（如 mermaid-filter.err）
            
        Returns:
            DOCX 文件的字节流
            
        Raises:
            PandocError: 转换失败时
        """
        if Keep_original_formula:
            md = self._convert_html_to_md(html_text, custom_filters_html_to_md)
            return self.convert_to_docx_bytes(
                    md_text=md,
                    reference_docx=reference_docx,
                    Keep_original_formula=Keep_original_formula,
                    enable_latex_replacements=enable_latex_replacements,
                    custom_filters=custom_filters_md_to_docx or custom_filters,
                    request_headers=request_headers,
                    cwd=cwd,
                )
        
        cmd = [
            self.pandoc_path,
            "-f", "html+tex_math_dollars+raw_tex+tex_math_double_backslash+tex_math_single_backslash",
            "-t", "docx",
            "-o", "-",
            "--highlight-style", "tango",
        ]
        if enable_latex_replacements:
            cmd += ["--lua-filter", LUA_LATEX_REPLACEMENTS]
        # 添加自定义 Filter
        cmd += self._build_filter_args(custom_filters)
        if reference_docx:
            cmd += ["--reference-doc", reference_docx]
        cmd = _add_request_headers(cmd, request_headers)

        startupinfo = None
        creationflags = 0
        if os.name == "nt":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            creationflags = subprocess.CREATE_NO_WINDOW

        # 确保工作目录存在且可写
        if cwd:
            cwd = os.path.expandvars(cwd)
            os.makedirs(cwd, exist_ok=True)

        # 关键：input 直接传 UTF-8 字节；text=False 以得到二进制 stdout
        result = subprocess.run(
            cmd,
            input=html_text.encode("utf-8"),
            capture_output=True,
            text=False,
            shell=False,
            startupinfo=startupinfo,
            creationflags=creationflags,
            cwd=cwd,
        )
        if result.returncode != 0:
            # stderr 可能是字节，转成字符串便于日志查看
            err = (result.stderr or b"").decode("utf-8", "ignore")
            log(f"Pandoc HTML conversion error: {err}")
            raise PandocError(err or "Pandoc HTML conversion failed")

        _log_pandoc_stderr_as_warning(result.stderr, context="Pandoc warning (HTML->DOCX)")
        return result.stdout

    def _strip_latex_preamble(self, latex: str) -> str:
        """
        Remove LaTeX document structure, keeping only body content.
        
        Strips:
        - \\documentclass, \\usepackage, \\begin{document}, \\end{document}
        - \\maketitle, \\tightlist, other preamble commands
        - Empty lines at start/end
        """
        lines = latex.split('\n')
        result_lines = []
        in_document = False
        skip_patterns = [
            r'^\s*\\documentclass',
            r'^\s*\\usepackage',
            r'^\s*\\begin\{document\}',
            r'^\s*\\end\{document\}',
            r'^\s*\\maketitle',
            r'^\s*\\date\{',
            r'^\s*\\author\{',
            r'^\s*\\providecommand',
            r'^\s*\\setlength',
            r'^\s*\\def\\tightlist',
            r'^\s*\\tightlist',
            r'^\s*\\newcommand',
        ]
        
        for line in lines:
            # Skip preamble lines
            should_skip = False
            for pattern in skip_patterns:
                if re.match(pattern, line):
                    should_skip = True
                    break
            
            if should_skip:
                if '\\begin{document}' in line:
                    in_document = True
                continue
            
            # Only include content after \begin{document} or if no document structure
            if in_document or not any(re.match(r'^\s*\\documentclass', l) for l in lines[:20]):
                result_lines.append(line)
        
        # Clean up result
        result = '\n'.join(result_lines)
        # Remove leading/trailing empty lines
        result = result.strip()
        return result

    def convert_html_to_latex_text(
        self,
        html_text: str,
        *,
        strip_preamble: bool = True,
        custom_filters_html_to_md: Optional[List[str]] = None,
        custom_filters_md_to_latex: Optional[List[str]] = None,
    ) -> str:
        """
        Convert HTML to LaTeX text.
        
        Args:
            html_text: HTML content
            strip_preamble: If True, remove document preamble for Overleaf paste
            
        Returns:
            LaTeX content (body only if strip_preamble=True)
        """
        # First convert HTML to Markdown
        md_text = self._convert_html_to_md(html_text, custom_filters_html_to_md)
        # Then convert Markdown to LaTeX
        return self.convert_markdown_to_latex_text(
            md_text,
            strip_preamble=strip_preamble,
            custom_filters=custom_filters_md_to_latex,
        )

    def convert_markdown_to_latex_text(
        self,
        md_text: str,
        *,
        strip_preamble: bool = True,
        enable_latex_replacements: bool = True,
        custom_filters: Optional[List[str]] = None,
    ) -> str:
        """
        Convert Markdown to LaTeX text.
        
        Args:
            md_text: Markdown content
            strip_preamble: If True, remove document preamble for Overleaf paste
            enable_latex_replacements: Enable LaTeX syntax fixes
            custom_filters: Optional list of custom Pandoc filters
            
        Returns:
            LaTeX content (body only if strip_preamble=True)
        """
        cmd = [
            self.pandoc_path,
            "-f", "markdown+tex_math_dollars+raw_tex+tex_math_double_backslash+tex_math_single_backslash",
            "-t", "latex",
            "-o", "-",
            "--wrap", "none",
        ]
        if enable_latex_replacements:
            cmd += ["--lua-filter", LUA_LATEX_REPLACEMENTS]
        cmd += self._build_filter_args(custom_filters)

        startupinfo = None
        creationflags = 0
        if os.name == "nt":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            creationflags = subprocess.CREATE_NO_WINDOW

        result = subprocess.run(
            cmd,
            input=md_text.encode("utf-8"),
            capture_output=True,
            text=False,
            shell=False,
            startupinfo=startupinfo,
            creationflags=creationflags,
        )
        if result.returncode != 0:
            err = (result.stderr or b"").decode("utf-8", "ignore")
            log(f"Pandoc Markdown to LaTeX error: {err}")
            raise PandocError(err or "Pandoc Markdown to LaTeX conversion failed")

        latex = result.stdout.decode("utf-8", "ignore")
        
        if strip_preamble:
            latex = self._strip_latex_preamble(latex)
        
        return latex
