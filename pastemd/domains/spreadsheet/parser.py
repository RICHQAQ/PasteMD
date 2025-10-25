"""Markdown table parser."""

import re
from typing import List, Optional


def _split_table_cells(line: str) -> List[str]:
    """
    按 | 分割表格单元格,正确处理转义的竖线
    
    Args:
        line: 表格行文本
        
    Returns:
        单元格列表
    """
    cells = []
    current_cell = []
    i = 0
    
    while i < len(line):
        if i > 0 and line[i] == '|' and line[i - 1] == '\\':
            # 转义的竖线,替换前一个反斜杠并添加竖线
            current_cell[-1] = '|'
            i += 1
        elif line[i] == '|':
            # 未转义的竖线,分割单元格
            cells.append(''.join(current_cell).strip())
            current_cell = []
            i += 1
        else:
            current_cell.append(line[i])
            i += 1
    
    # 添加最后一个单元格
    if current_cell or cells:  # 如果有内容或已经有单元格
        cells.append(''.join(current_cell).strip())
    
    return cells


def parse_markdown_table(md_text: str) -> Optional[List[List[str]]]:
    """
    解析 Markdown 表格为二维数组
    
    Args:
        md_text: Markdown 文本内容
        
    Returns:
        二维数组，每个元素代表一行的单元格内容；如果不是表格则返回 None
    """
    lines = md_text.strip().split('\n')
    if len(lines) < 2:
        return None
    
    table_data = []
    separator_found = False
    
    for i, line in enumerate(lines):
        line = line.strip()
        
        # 跳过空行
        if not line:
            continue
            
        # 检查是否为表格行（以 | 开头或结尾）
        if not (line.startswith('|') or line.endswith('|') or '|' in line):
            # 如果已经找到分隔符，说明表格结束
            if separator_found:
                break
            # 否则不是表格
            return None
        
        # 检查是否为分隔符行（如 |---|---|）
        if re.match(r'^\s*\|?\s*[-:]+\s*(\|\s*[-:]+\s*)+\|?\s*$', line):
            separator_found = True
            continue
        
        # 使用新的分割方法解析单元格
        cells = _split_table_cells(line)
        
        # 移除首尾的空元素（如果行是 |a|b| 格式）
        if cells and cells[0] == '':
            cells = cells[1:]
        if cells and cells[-1] == '':
            cells = cells[:-1]
        
        if cells:
            table_data.append(cells)
    
    # 必须找到分隔符才认为是有效表格
    if not separator_found or not table_data:
        return None
    
    return table_data
