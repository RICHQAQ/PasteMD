# -*- coding: utf-8 -*-
"""OMML (Office MathML) utilities for Office formula paste.

This module provides utilities for converting MathML to OMML and generating
Office-compatible HTML with formula support.
"""

from __future__ import annotations

import re
from html.entities import name2codepoint
from typing import Callable

from .logging import log


def convert_mathml_to_omml(mathml: str, entity_map: dict | None = None) -> str:
    """Convert MathML to OMML using mathml2omml library.
    
    Args:
        mathml: MathML XML string
        entity_map: Optional entity name to codepoint mapping (default: html.entities.name2codepoint)
        
    Returns:
        OMML XML string
        
    Raises:
        ImportError: If mathml2omml library is not installed
        ValueError: If conversion fails
    """
    try:
        import mathml2omml
    except ImportError:
        raise ImportError("mathml2omml library is required. Install with: pip install mathml2omml")
    
    if entity_map is None:
        entity_map = name2codepoint
    
    try:
        return mathml2omml.convert(mathml, entity_map)
    except Exception as e:
        raise ValueError(f"Failed to convert MathML to OMML: {e}")


def extract_mathml_elements(html: str) -> list[tuple[str, int, int]]:
    """Extract MathML elements from HTML.
    
    Args:
        html: HTML string containing MathML elements
        
    Returns:
        List of (mathml_string, start_pos, end_pos) tuples
    """
    pattern = r'<math[^>]*>.*?</math>'
    matches = []
    for match in re.finditer(pattern, html, re.DOTALL | re.IGNORECASE):
        matches.append((match.group(0), match.start(), match.end()))
    return matches


def wrap_omml_conditional(omml: str, fallback_text: str = "") -> str:
    """Wrap OMML in Office conditional comments.
    
    Args:
        omml: OMML XML string
        fallback_text: Text to show in non-Office applications
        
    Returns:
        HTML with conditional comments for Office
    """
    result = f'<!--[if gte msEquation 12]>{omml}<![endif]-->'
    if fallback_text:
        result += f'<![if !msEquation]>{fallback_text}<![endif]>'
    return result


def _extract_table_ranges(html: str) -> list[tuple[int, int]]:
    """Extract table ranges from HTML.

    Args:
        html: HTML string containing table elements

    Returns:
        List of (start_pos, end_pos) tuples for table blocks
    """
    pattern = r'<table[^>]*>.*?</table>'
    ranges = []
    for match in re.finditer(pattern, html, re.DOTALL | re.IGNORECASE):
        ranges.append((match.start(), match.end()))
    return ranges


def _is_within_ranges(start: int, end: int, ranges: list[tuple[int, int]]) -> bool:
    for range_start, range_end in ranges:
        if start >= range_start and end <= range_end:
            return True
    return False


def convert_html_mathml_to_omml(html: str, *, skip_table_mathml: bool = False) -> str:
    """Replace MathML elements in HTML with OMML conditional comments.

    Args:
        html: HTML string with MathML formulas
        skip_table_mathml: When True, keep MathML inside <table> blocks unchanged

    Returns:
        HTML with MathML replaced by OMML conditional comments
    """
    mathml_elements = extract_mathml_elements(html)
    if not mathml_elements:
        return html

    table_ranges: list[tuple[int, int]] = []
    if skip_table_mathml:
        table_ranges = _extract_table_ranges(html)

    # Process in reverse order to preserve positions
    result = html
    for mathml, start, end in reversed(mathml_elements):
        if table_ranges and _is_within_ranges(start, end, table_ranges):
            continue
        try:
            omml = convert_mathml_to_omml(mathml)
            # Extract original text content as fallback
            fallback = re.sub(r'<[^>]+>', '', mathml)
            replacement = wrap_omml_conditional(omml, fallback)
            result = result[:start] + replacement + result[end:]
        except Exception as e:
            log(f"Failed to convert MathML element: {e}")
            # Keep original MathML if conversion fails
            continue

    return result


def generate_office_html(
    body_content: str,
    *,
    title: str = "",
    lang: str = "zh-CN",
    font_family: str = "Calibri",
    font_size: str = "11.0pt",
) -> str:
    """Generate Office-compatible HTML with proper namespaces and metadata.
    
    Args:
        body_content: HTML content for the body
        title: Optional document title
        lang: Language code
        font_family: Default font family
        font_size: Default font size
        
    Returns:
        Complete Office HTML document
    """
    return f'''<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="Generator" content="PasteMD">
</head>
<body lang="{lang}" style="font-family:{font_family};font-size:{font_size}">
<!--StartFragment-->
{body_content}
<!--EndFragment-->
</body>
</html>'''


__all__ = [
    "convert_mathml_to_omml",
    "extract_mathml_elements",
    "wrap_omml_conditional",
    "convert_html_mathml_to_omml",
    "generate_office_html",
]
