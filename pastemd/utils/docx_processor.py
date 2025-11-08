"""DOCX document post-processing utilities."""

import io
from docx import Document
from ..utils.logging import log


class DocxProcessor:
    """DOCX 文档后处理器 - 用于修改已生成的 DOCX 文档样式"""
    
    @staticmethod
    def normalize_first_paragraph_style(
        docx_bytes: bytes,
        target_style: str = "Body Text"
    ) -> bytes:
        """
        将 DOCX 文档中的 "First Paragraph" 样式替换为指定样式
        
        Args:
            docx_bytes: DOCX 文件的字节流
            target_style: 目标样式名称，默认为 "Body Text"
            
        Returns:
            修改后的 DOCX 文件字节流
        """
        try:
            # 从字节流加载文档
            doc = Document(io.BytesIO(docx_bytes))
            
            # 统计修改的段落数量
            modified_count = 0
            
            # 遍历所有段落
            for paragraph in doc.paragraphs:
                # 检查段落样式是否为 "First Paragraph"
                if paragraph.style and paragraph.style.name == "First Paragraph":
                    # 修改为目标样式
                    paragraph.style = target_style
                    modified_count += 1
                    log(f"Changed paragraph style from 'First Paragraph' to '{target_style}'")
            
            # 如果有修改，记录日志
            if modified_count > 0:
                log(f"Total {modified_count} paragraph(s) changed from 'First Paragraph' to '{target_style}'")
            else:
                log("No 'First Paragraph' style found in document")
            
            # 将修改后的文档保存到字节流
            output_stream = io.BytesIO()
            doc.save(output_stream)
            output_stream.seek(0)
            
            return output_stream.read()
            
        except Exception as e:
            log(f"Failed to process DOCX styles: {type(e).__name__}: {e}")
            # 如果处理失败，返回原始字节流
            return docx_bytes
    
    @staticmethod
    def apply_custom_processing(
        docx_bytes: bytes,
        disable_first_para_indent: bool = False,
        target_style: str = "Body Text"
    ) -> bytes:
        """
        对 DOCX 文档应用自定义后处理
        
        Args:
            docx_bytes: DOCX 文件的字节流
            disable_first_para_indent: 是否禁用第一段特殊格式（替换 First Paragraph 样式）
            target_style: 目标样式名称
            
        Returns:
            处理后的 DOCX 文件字节流
        """
        # 如果需要禁用第一段特殊格式
        if disable_first_para_indent:
            docx_bytes = DocxProcessor.normalize_first_paragraph_style(
                docx_bytes,
                target_style
            )
        
        # 可以在这里添加其他后处理逻辑
        
        return docx_bytes
