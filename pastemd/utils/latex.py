"""LaTeX formula conversion utilities."""

import re


def convert_latex_delimiters(text: str) -> str:
    """
    弃用
    将 LaTeX 块级公式格式 \\[...\\] 转换为 Pandoc 支持的 $$...$$ 格式
    将 LaTeX 行内公式格式 \\(...\\) 转换为 Pandoc 支持的 $...$ 格式
    
    Args:
        text: 包含 LaTeX 公式的文本
        
    Returns:
        转换后的文本
    """
    # # 匹配 \[ 开始到 \] 结束的公式块
    # pattern = r'\\\[(.*?)\\\]'
    # inline_pattern = r'\\\((.*?)\\\)'

    # def replace_match(match):
    #     formula = match.group(1).strip()
    #     return f"$$\n{formula}\n$$"

    # def replace_inline_match(match):
    #     formula = match.group(1).strip()
    #     return f"${formula}$"

    # text = re.sub(pattern, replace_match, text, flags=re.DOTALL)
    # text = re.sub(inline_pattern, replace_inline_match, text, flags=re.DOTALL)
    return text
