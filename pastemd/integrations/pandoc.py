"""Pandoc CLI tool integration."""

import os
import subprocess
from typing import Optional

from ..core.errors import PandocError
from ..utils.logging import log


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
    
    def convert_to_docx(
        self,
        md_text: str,
        output_path: str,
        reference_docx: Optional[str] = None
    ) -> None:
        """
        将 Markdown 文本转换为 DOCX 文件

        Args:
            md_text: Markdown 文本内容
            output_path: 输出 DOCX 文件路径
            reference_docx: 可选的参考文档模板路径

        Raises:
            PandocError: 转换失败时
        """

        # 构建 Pandoc 命令
        cmd = [
            self.pandoc_path,
            "--from", "markdown+tex_math_dollars+raw_tex",
            "--to", "docx",
            "-o", output_path,
            "--highlight-style", "tango"
        ]

        if reference_docx:
            cmd.extend(["--reference-doc", reference_docx])

        try:
            # 在 Windows 上隐藏控制台窗口
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
                error_msg = result.stderr.strip() or result.stdout or "Pandoc conversion failed"
                log(f"Pandoc error: {error_msg}")
                raise PandocError(error_msg)

        except FileNotFoundError:
            raise PandocError(f"Pandoc executable not found: {self.pandoc_path}")
        except Exception as e:
            log(f"Pandoc conversion failed: {e}")
            raise PandocError(f"Conversion failed: {e}")

    def convert_to_docx_bytes(self, md_text: str, reference_docx: Optional[str] = None) -> bytes:
        """
        用 stdin 喂入 Markdown，直接把 DOCX 从 stdout 读到内存（无任何输入文件写盘）
        """
        cmd = [
            self.pandoc_path,
            "-f", "markdown+tex_math_dollars+raw_tex+tex_math_double_backslash+tex_math_single_backslash",
            "-t", "docx",
            "-o", "-",
            "--highlight-style", "tango",
        ]
        if reference_docx:
            cmd += ["--reference-doc", reference_docx]

        startupinfo = None
        creationflags = 0
        if os.name == "nt":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            creationflags = subprocess.CREATE_NO_WINDOW

        # 关键：input 直接传 UTF-8 字节；text=False 以得到二进制 stdout
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
            # stderr 可能是字节，转成字符串便于日志查看
            err = (result.stderr or b"").decode("utf-8", "ignore")
            log(f"Pandoc error: {err}")
            raise PandocError(err or "Pandoc conversion failed")

        return result.stdout

    def convert_html_to_docx_bytes(self, html_text: str, reference_docx: Optional[str] = None) -> bytes:
        """
        用 stdin 喂入 HTML，直接把 DOCX 从 stdout 读到内存（无任何输入文件写盘）
        
        Args:
            html_text: HTML 文本内容
            reference_docx: 可选的参考文档模板路径
            
        Returns:
            DOCX 文件的字节流
            
        Raises:
            PandocError: 转换失败时
        """
        cmd = [
            self.pandoc_path,
            "-f", "html+tex_math_dollars+raw_tex+tex_math_double_backslash+tex_math_single_backslash",
            "-t", "docx",
            "-o", "-",
            "--highlight-style", "tango",
        ]
        if reference_docx:
            cmd += ["--reference-doc", reference_docx]

        startupinfo = None
        creationflags = 0
        if os.name == "nt":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            creationflags = subprocess.CREATE_NO_WINDOW

        # 关键：input 直接传 UTF-8 字节；text=False 以得到二进制 stdout
        result = subprocess.run(
            cmd,
            input=html_text.encode("utf-8"),
            capture_output=True,
            text=False,
            shell=False,
            startupinfo=startupinfo,
            creationflags=creationflags,
        )
        if result.returncode != 0:
            # stderr 可能是字节，转成字符串便于日志查看
            err = (result.stderr or b"").decode("utf-8", "ignore")
            log(f"Pandoc HTML conversion error: {err}")
            raise PandocError(err or "Pandoc HTML conversion failed")

        return result.stdout
