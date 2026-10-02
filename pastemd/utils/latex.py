"""LaTeX formula conversion utilities."""

import re


def convert_latex_delimiters(text: str, fix_single_dollar_block: bool = True) -> str:
    """
    预处理 LaTeX 公式格式，使其能被 Pandoc 正确识别
    
    Args:
        text: 原始 Markdown 文本
        fix_single_dollar_block: 是否启用非标准块级公式修复（含单行 $ 块和行内 $ 空格）
        
    Returns:
        转换后的文本
    """
    # 1. (弃用) 标准 LaTeX 分隔符转换（\[...\] -> $$...$$）
    # 目前此功能被注释掉，若启用需解开 _convert_standard_latex_delimiters 内部注释
    text = _convert_standard_latex_delimiters(text)
    
    if fix_single_dollar_block:
        # 2. 修复行内公式中 $ 两侧的多余空格 ($  L  $ -> $L$)
        text = _fix_inline_math_spaces(text)
        
        # 3. 将单独一行的 $ ... $ 块级公式转换为 $$ ... $$
        text = _fix_single_dollar_blocks(text)
        
    return text


def _convert_standard_latex_delimiters(text: str) -> str:
    """
    (弃用) 将 LaTeX 标准分隔符转换为 Pandoc 支持的格式
    \\[ ... \\] -> $$ ... $$
    \\( ... \\) -> $ ... $
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


_CJK_CHARS_RE = re.compile(r'[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]')
_TEX_COMMAND_RE = re.compile(r'\\[a-zA-Z]+')


def _fix_inline_math_spaces(text: str) -> str:
    """
    修复行内公式中 $ 后面的空格和 $ 前面的空格

    Pandoc tex_math_dollars 要求 $ 后不能有空格，$ 前不能有空格
    例如：$  L  $ -> $L$、$ {x}$ -> ${x}$、$x $ -> $x$

    按行内 $ 的出现顺序逐对配对（与 Pandoc 的行内公式解析一致），
    避免正则方案把两个公式之间的普通文字（如 $a$ 和 $b$ 中的 " 和 "）
    误配成一个"公式"。围栏代码块（``` / ~~~）内的内容不做处理。
    """
    if '$' not in text:
        return text

    out = []
    in_code = False
    fence = ""
    for line in text.split('\n'):
        stripped = line.strip()
        if stripped.startswith('```') or stripped.startswith('~~~'):
            if not in_code:
                in_code, fence = True, stripped[:3]
            elif stripped.startswith(fence):
                in_code, fence = False, ""
            out.append(line)
            continue
        out.append(line if in_code else _fix_inline_spaces_in_line(line))
    return '\n'.join(out)


def _fix_inline_spaces_in_line(line: str) -> str:
    """对单行执行 $...$ 配对与内侧空格修复。"""
    if '$' not in line:
        return line

    out = []
    i = 0
    n = len(line)
    while i < n:
        ch = line[i]
        if ch != '$':
            out.append(ch)
            i += 1
            continue

        # 转义的美元符号 \$：普通字符
        if i > 0 and line[i - 1] == '\\':
            out.append(ch)
            i += 1
            continue

        # $$：块级公式定界符或转义，原样输出
        if i + 1 < n and line[i + 1] == '$':
            out.append('$$')
            i += 2
            continue

        # 找与之配对的闭合 $
        close = line.find('$', i + 1)
        if close == -1:
            out.append(line[i:])
            break

        content = line[i + 1:close]
        stripped = content.strip()
        # 空内容不当作公式；含中日韩字符视为普通文本（如 "$ 5 和 $"），
        # 但含 LaTeX 命令的除外（如 "$ \text{面积} = 5 $" 是真公式）
        if not stripped or (
            _CJK_CHARS_RE.search(stripped) and not _TEX_COMMAND_RE.search(stripped)
        ):
            out.append('$')
            i += 1
            continue

        if stripped != content:
            out.append(f'${stripped}$')
        else:
            out.append(line[i:close + 1])
        i = close + 1

    return ''.join(out)


def _fix_single_dollar_blocks(text: str) -> str:
    """
    将单独一行的 $ ... $ 块级公式转换为 $$ ... $$
    规则：如果某一行去除首尾空白后只有 $，则视为块公式的分隔符
    注意要跳过代码块
    """
    lines = text.split('\n')
    new_lines = []
    in_code_block = False
    code_fence_char = ""
    in_dollar_block = False
    
    for line in lines:
        stripped = line.strip()
        
        # 1. 代码块检测
        if stripped.startswith('```') or stripped.startswith('~~~'):
            fence = stripped[:3]
            if not in_code_block:
                in_code_block = True
                code_fence_char = fence
                new_lines.append(line)
                continue
            elif stripped.startswith(code_fence_char):
                in_code_block = False
                code_fence_char = ""
                new_lines.append(line)
                continue
            
        if in_code_block:
            new_lines.append(line)
            continue
            
        # 2. 单行 $ 检测
        # 匹配仅包含 $ 的行（允许缩进）
        if re.match(r'^\s*\$\s*$', line):
            # 这是一个潜在的块公式分隔符
            if not in_dollar_block:
                # 块开始：替换为 $$
                # 尽量保留原有的缩进
                prefix = line[:line.find('$')]
                new_lines.append(f"{prefix}$$")
                in_dollar_block = True
            else:
                # 块结束：替换为 $$
                prefix = line[:line.find('$')]
                new_lines.append(f"{prefix}$$")
                in_dollar_block = False
        else:
            new_lines.append(line)
            
    return '\n'.join(new_lines)
