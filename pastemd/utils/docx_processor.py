"""DOCX document post-processing utilities."""

import io
import zipfile
from typing import Optional
from xml.sax.saxutils import quoteattr
from xml.etree import ElementTree as ET

from docx import Document
from ..utils.logging import log


_WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_MARKUP_COMPATIBILITY_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_DOCX_TABLE_WIDTH_TWIPS = 9360
_MIN_TABLE_COLUMN_TWIPS = 900
_MAX_LABEL_COLUMN_TWIPS = 2200


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
    def replace_horizontal_rules_with_paragraph_borders(docx_bytes: bytes) -> bytes:
        """
        将 Pandoc 生成的 VML 横线替换为 Word/WPS 更兼容的段落边框。

        Pandoc 的 Markdown HorizontalRule 在 DOCX 中会生成
        ``<w:pict><v:rect ... o:hr="t" /></w:pict>``。WPS 可能把这个结构显示
        成矩形块；段落边框更接近 Word/WPS 手动输入 ``---`` 后回车的结果。
        """
        namespaces = {
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
            "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
            "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
            "v": "urn:schemas-microsoft-com:vml",
            "o": "urn:schemas-microsoft-com:office:office",
            "w10": "urn:schemas-microsoft-com:office:word",
        }
        for prefix, uri in namespaces.items():
            ET.register_namespace(prefix, uri)

        document_path = "word/document.xml"

        try:
            input_stream = io.BytesIO(docx_bytes)
            output_stream = io.BytesIO()

            with zipfile.ZipFile(input_stream, "r") as zin:
                document_xml = zin.read(document_path)
                xml_namespaces = DocxProcessor._extract_xml_namespaces(document_xml)
                DocxProcessor._register_xml_namespaces(xml_namespaces)
                root = ET.fromstring(document_xml)

                modified_count = 0
                for paragraph in root.findall(".//w:p", namespaces):
                    rects = paragraph.findall(".//v:rect", namespaces)
                    if not any(rect.get(f"{{{namespaces['o']}}}hr") == "t" for rect in rects):
                        continue

                    paragraph.clear()
                    ppr = ET.SubElement(paragraph, f"{{{namespaces['w']}}}pPr")
                    spacing = ET.SubElement(ppr, f"{{{namespaces['w']}}}spacing")
                    spacing.set(f"{{{namespaces['w']}}}before", "0")
                    spacing.set(f"{{{namespaces['w']}}}after", "0")
                    pbdr = ET.SubElement(ppr, f"{{{namespaces['w']}}}pBdr")
                    bottom = ET.SubElement(pbdr, f"{{{namespaces['w']}}}bottom")
                    bottom.set(f"{{{namespaces['w']}}}val", "single")
                    bottom.set(f"{{{namespaces['w']}}}sz", "6")
                    bottom.set(f"{{{namespaces['w']}}}space", "1")
                    bottom.set(f"{{{namespaces['w']}}}color", "auto")
                    modified_count += 1

                if modified_count == 0:
                    log("No VML horizontal rule found in DOCX")
                    return docx_bytes

                updated_document_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                updated_document_xml = DocxProcessor._preserve_ignorable_namespaces(
                    updated_document_xml,
                    root,
                    xml_namespaces,
                )

                with zipfile.ZipFile(output_stream, "w") as zout:
                    for item in zin.infolist():
                        data = updated_document_xml if item.filename == document_path else zin.read(item.filename)
                        zout.writestr(item, data)

            output_stream.seek(0)
            log(f"Replaced {modified_count} DOCX horizontal rule(s) with paragraph border(s)")
            return output_stream.read()

        except Exception as e:
            log(f"Failed to replace DOCX horizontal rules: {type(e).__name__}: {e}")
            return docx_bytes

    @staticmethod
    def auto_layout_tables(docx_bytes: bytes) -> bytes:
        """
        根据单元格文本长度调整 DOCX 普通表格列宽。

        该处理主要改善「左侧短标签、右侧长内容」的两列表格，让 Word/WPS
        打开后右侧内容列获得更多空间，减少不必要换行。
        """
        namespaces = {
            "w": _WORD_NS,
        }
        for prefix, uri in namespaces.items():
            ET.register_namespace(prefix, uri)

        document_path = "word/document.xml"

        try:
            input_stream = io.BytesIO(docx_bytes)
            output_stream = io.BytesIO()

            with zipfile.ZipFile(input_stream, "r") as zin:
                document_xml = zin.read(document_path)
                xml_namespaces = DocxProcessor._extract_xml_namespaces(document_xml)
                DocxProcessor._register_xml_namespaces(xml_namespaces)
                root = ET.fromstring(document_xml)

                modified_count = DocxProcessor._auto_layout_tables_in_element(
                    root,
                    namespaces,
                    _DOCX_TABLE_WIDTH_TWIPS,
                )

                if modified_count == 0:
                    log("No eligible DOCX table found for auto layout")
                    return docx_bytes

                updated_document_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                updated_document_xml = DocxProcessor._preserve_ignorable_namespaces(
                    updated_document_xml,
                    root,
                    xml_namespaces,
                )

                with zipfile.ZipFile(output_stream, "w") as zout:
                    for item in zin.infolist():
                        data = updated_document_xml if item.filename == document_path else zin.read(item.filename)
                        zout.writestr(item, data)

            output_stream.seek(0)
            log(f"Auto-layout applied to {modified_count} DOCX table(s)")
            return output_stream.read()

        except Exception as e:
            log(f"Failed to auto-layout DOCX tables: {type(e).__name__}: {e}")
            return docx_bytes

    @staticmethod
    def _extract_xml_namespaces(xml_bytes: bytes) -> dict[str, str]:
        namespaces: dict[str, str] = {}
        try:
            for _event, namespace in ET.iterparse(
                io.BytesIO(xml_bytes),
                events=("start-ns",),
            ):
                prefix, uri = namespace
                if prefix and prefix not in namespaces:
                    namespaces[prefix] = uri
        except Exception as e:
            log(f"Failed to extract DOCX XML namespaces: {type(e).__name__}: {e}")
        return namespaces

    @staticmethod
    def _register_xml_namespaces(namespaces: dict[str, str]) -> None:
        for prefix, uri in namespaces.items():
            if prefix in ("xml", "xmlns"):
                continue
            try:
                ET.register_namespace(prefix, uri)
            except ValueError as e:
                log(f"Failed to register DOCX XML namespace {prefix}: {e}")

    @staticmethod
    def _preserve_ignorable_namespaces(
        xml_bytes: bytes,
        root: ET.Element,
        namespaces: dict[str, str],
    ) -> bytes:
        ignorable = root.get(f"{{{_MARKUP_COMPATIBILITY_NS}}}Ignorable")
        if not ignorable:
            return xml_bytes

        xml_text = xml_bytes.decode("utf-8")
        root_start = xml_text.find(":document")
        if root_start < 0:
            return xml_bytes
        root_start = xml_text.rfind("<", 0, root_start)
        if root_start < 0:
            return xml_bytes
        root_end = xml_text.find(">", root_start)
        if root_end < 0:
            return xml_bytes

        root_tag = xml_text[root_start:root_end]
        additions = []
        for prefix in ignorable.split():
            uri = namespaces.get(prefix)
            if uri and f"xmlns:{prefix}=" not in root_tag:
                additions.append(f" xmlns:{prefix}={quoteattr(uri)}")

        if not additions:
            return xml_bytes

        xml_text = xml_text[:root_end] + "".join(additions) + xml_text[root_end:]
        return xml_text.encode("utf-8")

    @staticmethod
    def _auto_layout_tables_in_element(
        element: ET.Element,
        namespaces: dict,
        available_width: int,
    ) -> int:
        w = f"{{{_WORD_NS}}}"
        modified_count = 0

        for child in list(element):
            if child.tag == f"{w}tbl":
                modified_count += DocxProcessor._auto_layout_table_with_nested(
                    child,
                    namespaces,
                    available_width,
                )
            else:
                modified_count += DocxProcessor._auto_layout_tables_in_element(
                    child,
                    namespaces,
                    available_width,
                )

        return modified_count

    @staticmethod
    def _auto_layout_table_with_nested(
        table: ET.Element,
        namespaces: dict,
        available_width: int,
    ) -> int:
        modified_count = 0
        table_width = DocxProcessor._available_table_width(
            table,
            namespaces,
            available_width,
        )
        widths = DocxProcessor._suggest_table_column_widths(
            table,
            namespaces,
            table_width,
        )

        if widths:
            DocxProcessor._apply_table_column_widths(table, widths, namespaces)
            modified_count += 1
        else:
            widths = DocxProcessor._existing_table_column_widths(table, namespaces)

        for row in table.findall("./w:tr", namespaces):
            cells = row.findall("./w:tc", namespaces)
            for index, cell in enumerate(cells):
                cell_width = DocxProcessor._cell_width(cell, namespaces)
                if cell_width is None and index < len(widths):
                    cell_width = widths[index]
                if cell_width is None:
                    cell_width = table_width
                modified_count += DocxProcessor._auto_layout_tables_in_element(
                    cell,
                    namespaces,
                    cell_width,
                )

        return modified_count

    @staticmethod
    def _suggest_table_column_widths(
        table: ET.Element,
        namespaces: dict,
        total_width: int,
    ) -> list[int]:
        rows = table.findall("./w:tr", namespaces)
        if not rows:
            return []

        row_cells = [row.findall("./w:tc", namespaces) for row in rows]
        column_count = max((len(cells) for cells in row_cells), default=0)
        if column_count < 2:
            return []
        if total_width < column_count:
            return []

        if DocxProcessor._has_merged_cells(row_cells, namespaces):
            return []

        scores = [0.0] * column_count
        for cells in row_cells:
            for index, cell in enumerate(cells):
                text = DocxProcessor._direct_cell_text(cell)
                scores[index] = max(scores[index], DocxProcessor._visual_text_length(text))

        if not any(scores):
            return []

        return DocxProcessor._column_widths_from_scores(scores, total_width)

    @staticmethod
    def _available_table_width(
        table: ET.Element,
        namespaces: dict,
        fallback_width: int,
    ) -> int:
        tbl_width = table.find("./w:tblPr/w:tblW", namespaces)
        if tbl_width is not None:
            parsed_width = DocxProcessor._parse_twips(tbl_width.get(f"{{{_WORD_NS}}}w"))
            if parsed_width is not None:
                width_type = tbl_width.get(f"{{{_WORD_NS}}}type")
                if width_type == "dxa":
                    return parsed_width
                if width_type == "pct":
                    return max(1, int(fallback_width * parsed_width / 5000))

        tbl_layout = table.find("./w:tblPr/w:tblLayout", namespaces)
        if tbl_layout is not None and tbl_layout.get(f"{{{_WORD_NS}}}type") == "fixed":
            grid_widths = DocxProcessor._existing_table_column_widths(table, namespaces)
            if grid_widths:
                return sum(grid_widths)

        return fallback_width

    @staticmethod
    def _existing_table_column_widths(table: ET.Element, namespaces: dict) -> list[int]:
        widths: list[int] = []
        for grid_col in table.findall("./w:tblGrid/w:gridCol", namespaces):
            width = DocxProcessor._parse_twips(grid_col.get(f"{{{_WORD_NS}}}w"))
            if width is None:
                return []
            widths.append(width)
        return widths

    @staticmethod
    def _cell_width(cell: ET.Element, namespaces: dict) -> Optional[int]:
        tc_width = cell.find("./w:tcPr/w:tcW", namespaces)
        if tc_width is None:
            return None
        if tc_width.get(f"{{{_WORD_NS}}}type") not in (None, "dxa"):
            return None
        return DocxProcessor._parse_twips(tc_width.get(f"{{{_WORD_NS}}}w"))

    @staticmethod
    def _parse_twips(value: Optional[str]) -> Optional[int]:
        if value is None:
            return None
        try:
            width = int(value)
        except ValueError:
            return None
        return width if width > 0 else None

    @staticmethod
    def _direct_cell_text(cell: ET.Element) -> str:
        w = f"{{{_WORD_NS}}}"
        parts: list[str] = []

        def walk(element: ET.Element, inside_nested_table: bool) -> None:
            for child in list(element):
                child_inside_nested_table = inside_nested_table or child.tag == f"{w}tbl"
                if child.tag == f"{w}t" and not child_inside_nested_table:
                    parts.append(child.text or "")
                walk(child, child_inside_nested_table)

        walk(cell, False)
        return "".join(parts)

    @staticmethod
    def _has_merged_cells(row_cells: list[list[ET.Element]], namespaces: dict) -> bool:
        for cells in row_cells:
            for cell in cells:
                grid_span = cell.find("./w:tcPr/w:gridSpan", namespaces)
                if grid_span is not None and grid_span.get(f"{{{_WORD_NS}}}val") not in (None, "1"):
                    return True
                if cell.find("./w:tcPr/w:vMerge", namespaces) is not None:
                    return True
        return False

    @staticmethod
    def _visual_text_length(text: str) -> float:
        length = 0.0
        for char in text:
            if char.isspace():
                length += 0.3
            elif ord(char) > 127:
                length += 2.0
            else:
                length += 1.0
        return length

    @staticmethod
    def _column_widths_from_scores(scores: list[float], total_width: int) -> list[int]:
        weights = [max(score, 4.0) for score in scores]

        if len(weights) == 2 and weights[0] * 1.5 < weights[1]:
            first_share = weights[0] / sum(weights)
            first_width = int(total_width * max(0.14, min(first_share * 1.15, 0.24)))
            if total_width >= _MIN_TABLE_COLUMN_TWIPS * 2:
                first_width = max(
                    _MIN_TABLE_COLUMN_TWIPS,
                    min(
                        first_width,
                        _MAX_LABEL_COLUMN_TWIPS,
                        total_width - _MIN_TABLE_COLUMN_TWIPS,
                    ),
                )
            else:
                first_width = max(1, min(first_width, total_width - 1))
            return [first_width, total_width - first_width]

        if len(weights) * _MIN_TABLE_COLUMN_TWIPS > total_width:
            base_width = max(1, total_width // (len(weights) * 2))
        else:
            base_width = _MIN_TABLE_COLUMN_TWIPS

        widths = [base_width] * len(weights)
        remaining_width = total_width - sum(widths)
        if remaining_width < 0:
            base_width = max(1, total_width // len(weights))
            widths = [base_width] * len(weights)
            remaining_width = total_width - sum(widths)

        if remaining_width > 0:
            weight_sum = sum(weights)
            widths = [
                width + int(remaining_width * weight / weight_sum)
                for width, weight in zip(widths, weights)
            ]

        delta = total_width - sum(widths)
        if delta:
            index = max(range(len(widths)), key=lambda i: weights[i])
            widths[index] += delta
        return widths

    @staticmethod
    def _apply_table_column_widths(table: ET.Element, widths: list[int], namespaces: dict) -> None:
        w = f"{{{_WORD_NS}}}"

        tbl_pr = table.find("./w:tblPr", namespaces)
        if tbl_pr is None:
            tbl_pr = ET.Element(f"{w}tblPr")
            table.insert(0, tbl_pr)

        tbl_w = DocxProcessor._ensure_table_width_element(tbl_pr)
        if tbl_w.get(f"{w}type") != "pct":
            tbl_w.set(f"{w}type", "dxa")
            tbl_w.set(f"{w}w", str(sum(widths)))

        tbl_layout = DocxProcessor._ensure_table_layout_element(tbl_pr)
        tbl_layout.set(f"{w}type", "fixed")

        for child in list(table):
            if child.tag == f"{w}tblGrid":
                table.remove(child)

        tbl_grid = ET.Element(f"{w}tblGrid")
        for width in widths:
            grid_col = ET.SubElement(tbl_grid, f"{w}gridCol")
            grid_col.set(f"{w}w", str(width))
        table.insert(1 if table[0].tag == f"{w}tblPr" else 0, tbl_grid)

        for row in table.findall("./w:tr", namespaces):
            cells = row.findall("./w:tc", namespaces)
            for index, cell in enumerate(cells):
                if index >= len(widths):
                    continue
                tc_pr = cell.find("./w:tcPr", namespaces)
                if tc_pr is None:
                    tc_pr = ET.Element(f"{w}tcPr")
                    cell.insert(0, tc_pr)
                tc_w = DocxProcessor._ensure_cell_width_element(tc_pr)
                tc_w.set(f"{w}type", "dxa")
                tc_w.set(f"{w}w", str(widths[index]))

    @staticmethod
    def _ensure_table_width_element(tbl_pr: ET.Element) -> ET.Element:
        w = f"{{{_WORD_NS}}}"
        tbl_w = tbl_pr.find(f"./{w}tblW")
        if tbl_w is not None:
            return tbl_w

        tbl_w = ET.Element(f"{w}tblW")
        following_tags = {
            f"{w}jc",
            f"{w}tblCellSpacing",
            f"{w}tblInd",
            f"{w}tblBorders",
            f"{w}shd",
            f"{w}tblLayout",
            f"{w}tblCellMar",
            f"{w}tblLook",
            f"{w}tblCaption",
            f"{w}tblDescription",
            f"{w}tblPrChange",
        }
        insert_index = len(tbl_pr)
        for index, child in enumerate(list(tbl_pr)):
            if child.tag in following_tags:
                insert_index = index
                break
        tbl_pr.insert(insert_index, tbl_w)
        return tbl_w

    @staticmethod
    def _ensure_table_layout_element(tbl_pr: ET.Element) -> ET.Element:
        w = f"{{{_WORD_NS}}}"
        tbl_layout = tbl_pr.find(f"./{w}tblLayout")
        if tbl_layout is not None:
            tbl_pr.remove(tbl_layout)
        else:
            tbl_layout = ET.Element(f"{w}tblLayout")

        following_tags = {
            f"{w}tblCellMar",
            f"{w}tblLook",
            f"{w}tblCaption",
            f"{w}tblDescription",
            f"{w}tblPrChange",
        }
        insert_index = len(tbl_pr)
        for index, child in enumerate(list(tbl_pr)):
            if child.tag in following_tags:
                insert_index = index
                break
        tbl_pr.insert(insert_index, tbl_layout)
        return tbl_layout

    @staticmethod
    def _ensure_cell_width_element(tc_pr: ET.Element) -> ET.Element:
        w = f"{{{_WORD_NS}}}"
        tc_w = tc_pr.find(f"./{w}tcW")
        if tc_w is not None:
            return tc_w

        tc_w = ET.Element(f"{w}tcW")
        following_tags = {
            f"{w}gridSpan",
            f"{w}hMerge",
            f"{w}vMerge",
            f"{w}tcBorders",
            f"{w}shd",
            f"{w}noWrap",
            f"{w}tcMar",
            f"{w}textDirection",
            f"{w}tcFitText",
            f"{w}vAlign",
            f"{w}hideMark",
            f"{w}cellIns",
            f"{w}cellDel",
            f"{w}cellMerge",
            f"{w}tcPrChange",
        }
        insert_index = len(tc_pr)
        for index, child in enumerate(list(tc_pr)):
            if child.tag in following_tags:
                insert_index = index
                break
        tc_pr.insert(insert_index, tc_w)
        return tc_w
     
    @staticmethod
    def apply_custom_processing(
        docx_bytes: bytes,
        disable_first_para_indent: bool = False,
        target_style: str = "Body Text",
        horizontal_rule_style: str = "default",
        auto_layout_tables: bool = False,
    ) -> bytes:
        """
        对 DOCX 文档应用自定义后处理
        
        Args:
            docx_bytes: DOCX 文件的字节流
            disable_first_para_indent: 是否禁用第一段特殊格式（替换 First Paragraph 样式）
            target_style: 目标样式名称
            horizontal_rule_style: 水平线样式，paragraph_border 时转为段落边框
            auto_layout_tables: 是否自动调整普通表格列宽
            
        Returns:
            处理后的 DOCX 文件字节流
        """
        # 如果需要禁用第一段特殊格式
        if disable_first_para_indent:
            docx_bytes = DocxProcessor.normalize_first_paragraph_style(
                docx_bytes,
                target_style
            )
        
        if horizontal_rule_style == "paragraph_border":
            docx_bytes = DocxProcessor.replace_horizontal_rules_with_paragraph_borders(
                docx_bytes
            )

        if auto_layout_tables:
            docx_bytes = DocxProcessor.auto_layout_tables(docx_bytes)
        
        return docx_bytes
