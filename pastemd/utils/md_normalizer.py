"""Markdown 格式规范化工具 - 处理不同来源的 Markdown 格式差异"""

import re


def normalize_markdown(md_text: str) -> str:
    """
    规范化 Markdown 文本格式，确保元素之间有适当的空行
    
    主要处理以下问题：
    1. 标题前后缺少空行（如智谱清言）
    2. 代码块前后缺少空行
    3. 列表、引用等块级元素前后缺少空行
    4. 表格前后缺少空行
    
    Args:
        md_text: 原始 Markdown 文本
        
    Returns:
        规范化后的 Markdown 文本
    """
    # 统一换行符为 \n
    text = md_text.replace('\r\n', '\n').replace('\r', '\n')
    lines = text.split('\n')
    
    result = []
    in_code_block = False
    in_table = False
    prev_line_type = 'start'  # start, empty, text, heading, code, table, list, quote, hr
    
    for i, line in enumerate(lines):
        current_type = _get_line_type(line, in_code_block, in_table)
        
        # 代码块状态切换
        if line.startswith('```'):
            in_code_block = not in_code_block
        
        # 表格状态检测
        if current_type == 'table':
            in_table = True
        elif in_table and current_type not in ('table', 'empty'):
            in_table = False
        
        # 决定是否需要在当前行前添加空行
        need_blank_before = _should_add_blank_line(prev_line_type, current_type)
        
        if need_blank_before and result and result[-1].strip():
            result.append('')
        
        result.append(line)
        
        # 决定是否需要在当前行后添加空行
        need_blank_after = _should_add_blank_after(current_type, i, lines)
        
        if need_blank_after and i + 1 < len(lines) and lines[i + 1].strip():
            result.append('')
        
        # 更新前一行类型
        prev_line_type = current_type if line.strip() else 'empty'
    
    text = '\n'.join(result)
    
    # 清理多余的连续空行（超过2个连续换行符压缩为2个）
    text = re.sub(r'\n{3,}', '\n\n', text)
    
    # 恢复原始换行符风格
    if '\r\n' in md_text:
        text = text.replace('\n', '\r\n')
    
    return text


def _get_line_type(line: str, in_code_block: bool, in_table: bool) -> str:
    """判断行的类型"""
    stripped = line.strip()
    
    if not stripped:
        return 'empty'
    
    if in_code_block:
        return 'code'
    
    # 代码块边界
    if line.startswith('```'):
        return 'code'
    
    # 标题
    if re.match(r'^#{1,6}\s+', line):
        return 'heading'
    
    # 表格
    if line.startswith('|') and line.endswith('|'):
        return 'table'
    
    # 分隔线
    if re.match(r'^[-*_]{3,}$', stripped):
        return 'hr'
    
    # 列表（无序）
    if re.match(r'^[-*+]\s', line):
        return 'list'
    
    # 列表（有序）
    if re.match(r'^\d+\.\s', line):
        return 'list'
    
    # 引用
    if line.startswith('>'):
        return 'quote'
    
    return 'text'


def _should_add_blank_line(prev_type: str, current_type: str) -> bool:
    """判断两种类型之间是否需要空行"""
    # 文档开头或前一行是空行，不需要
    if prev_type in ('start', 'empty'):
        return False
    
    # 当前行是空行，不需要
    if current_type == 'empty':
        return False
    
    # 标题前需要空行（除非前面是标题）
    if current_type == 'heading':
        return prev_type not in ('heading',)
    
    # 代码块前需要空行
    if current_type == 'code' and prev_type != 'code':
        return True
    
    # 表格前需要空行
    if current_type == 'table' and prev_type not in ('table',):
        return True
    
    # 列表前需要空行（除非前面是列表）
    if current_type == 'list' and prev_type not in ('list',):
        return True
    
    # 引用前需要空行
    if current_type == 'quote' and prev_type not in ('quote',):
        return True
    
    # 分隔线前需要空行
    if current_type == 'hr':
        return True
    
    return False


def _should_add_blank_after(current_type: str, index: int, lines: list) -> bool:
    """判断当前行后是否需要空行"""
    # 最后一行不需要
    if index >= len(lines) - 1:
        return False
    
    next_line = lines[index + 1].strip()
    
    # 下一行已经是空行，不需要
    if not next_line:
        return False
    
    # 标题后需要空行
    if current_type == 'heading':
        return True
    
    # 代码块结束后需要空行
    if current_type == 'code' and lines[index].startswith('```'):
        # 检查是结束标记
        in_code = False
        for i in range(index):
            if lines[i].startswith('```'):
                in_code = not in_code
        if not in_code:  # 结束标记
            return True
    
    # 分隔线后需要空行
    if current_type == 'hr':
        return True
    
    return False

