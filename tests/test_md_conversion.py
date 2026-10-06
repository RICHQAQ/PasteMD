# -*- coding: utf-8 -*-
"""Markdown 预处理回归测试：内嵌 HTML 表格转换、行内公式空格修复、LaTeX 旧命令改写"""

import shutil
import subprocess
from pathlib import Path

import pytest

from pastemd.utils.latex import convert_latex_delimiters
from pastemd.utils.md_html_tables import convert_embedded_html_tables

PROJECT_ROOT = Path(__file__).resolve().parent.parent
LUA_FILTER = PROJECT_ROOT / "pastemd" / "lua" / "latex-replacements.lua"

PANDOC = shutil.which("pandoc")


def test_table_after_paragraph_is_converted():
    # <table> 紧跟段落文字时，pandoc 的 markdown reader 不解析 raw HTML 表格，
    # DOCX 输出会把单元格拍平为独立段落
    md = '表1. 平均正确率。\n<table><tr><td>被试</td><td>FBCCA</td></tr><tr><td>S1</td><td>0.6417</td></tr></table>\n\n后文。'
    out = convert_embedded_html_tables(md)
    assert '| 被试 | FBCCA |' in out
    assert '| S1 | 0.6417 |' in out
    assert '\n\n| 被试' in out
    assert '| --- |' in out
    assert '<table' not in out
    assert '<td>' not in out


def test_table_inside_code_fence_untouched():
    md = '```\n<table><tr><td>a</td></tr></table>\n```'
    assert convert_embedded_html_tables(md) == md


def test_table_inside_html_comment_untouched():
    md = '<!-- <table><tr><td>x</td></tr></table> -->'
    assert convert_embedded_html_tables(md) == md


def test_table_inside_indented_code_block_untouched():
    md = '正文：\n\n    <table><tr><td>x</td></tr></table>\n'
    assert convert_embedded_html_tables(md) == md


def test_first_row_becomes_header_when_no_th():
    md = '<table><tr><td>h1</td><td>h2</td></tr><tr><td>a</td><td>b</td></tr></table>'
    out = convert_embedded_html_tables(md)
    assert out.find('| h1 | h2 |') < out.find('| --- |') < out.find('| a | b |')


def test_pipe_in_cell_is_escaped():
    md = '<table><tr><td>a|b</td></tr></table>'
    assert '\\|' in convert_embedded_html_tables(md)


def test_empty_separator_row_is_dropped():
    md = '<table><tr><td>a</td></tr><tr><td colspan="3"></td></tr><tr><td>b</td></tr></table>'
    out = convert_embedded_html_tables(md)
    assert '|  |' not in out
    assert '| a |' in out
    assert '| b |' in out


def test_br_in_cell_becomes_space():
    md = '<table><tr><td>行1<br>行2</td></tr></table>'
    assert '| 行1 行2 |' in convert_embedded_html_tables(md)


def test_multiline_html_table_is_converted():
    md = '<table>\n<tr><td>a</td><td>b</td></tr>\n<tr><td>1</td><td>2</td></tr>\n</table>'
    out = convert_embedded_html_tables(md)
    assert '| a | b |' in out
    assert '| 1 | 2 |' in out


def test_text_without_table_untouched():
    md = '# 标题\n\n正文 $x$ 文本。'
    assert convert_embedded_html_tables(md) == md


def test_prose_between_formulas_not_matched():
    # 旧正则实现会把 $a$ 和 $b$ 之间的 " 和 " 误配成 $和$
    text = '信号 $Z _ { 1 }$ 和 $Z _ { 2 }$ 的相关'
    assert convert_latex_delimiters(text, fix_single_dollar_block=True) == text


def test_chinese_sentence_between_formulas_preserved():
    text = '其中 $N _ { s }$ 表示采样点数，且 $i \\in A$。'
    assert convert_latex_delimiters(text, fix_single_dollar_block=True) == text


def test_leading_space_in_inline_math_is_fixed():
    text = '$ { \\boldsymbol { Z } } _ { 2 }$'
    assert convert_latex_delimiters(text, fix_single_dollar_block=True) == '${ \\boldsymbol { Z } } _ { 2 }$'


def test_trailing_space_in_inline_math_is_fixed():
    assert convert_latex_delimiters('$x=1 $', fix_single_dollar_block=True) == '$x=1$'


def test_both_sides_spaces_in_inline_math_is_fixed():
    assert convert_latex_delimiters('$ x=1 $', fix_single_dollar_block=True) == '$x=1$'


def test_single_line_display_math_untouched():
    text = '$$ x + y $$'
    assert convert_latex_delimiters(text, fix_single_dollar_block=True) == text


def test_currency_like_text_not_converted():
    text = '价格是 $ 5 和 $ 10 元'
    assert convert_latex_delimiters(text, fix_single_dollar_block=True) == text


def test_escaped_dollar_not_treated_as_delimiter():
    text = '价格 \\$5 and 成本 \\$10'
    assert convert_latex_delimiters(text, fix_single_dollar_block=True) == text


def test_mixed_inline_math_and_prose():
    assert convert_latex_delimiters('由 $ a $ 可得 $ b $，其中', fix_single_dollar_block=True) == '由 $a$ 可得 $b$，其中'


def test_multiline_display_math_untouched():
    text = '$$\nH _ { 0 } : D = 0\n$$\n正文 $ x $ 尾行'
    out = convert_latex_delimiters(text, fix_single_dollar_block=True)
    assert '$$\nH _ { 0 } : D = 0\n$$' in out
    assert '$x$' in out


def test_inline_math_inside_code_fence_untouched():
    text = '```\n$ x = 1 $\n```'
    assert convert_latex_delimiters(text, fix_single_dollar_block=True) == text


def test_cjk_math_with_latex_command_is_fixed():
    text = '已知 $ \\text{面积} = 5 $，求 $ r $。'
    out = convert_latex_delimiters(text, fix_single_dollar_block=True)
    assert '$\\text{面积} = 5$' in out
    assert '$r$' in out


def test_cjk_text_without_latex_command_not_converted():
    text = '收入 $ 5，支出 $ 6 元'
    assert convert_latex_delimiters(text, fix_single_dollar_block=True) == text


@pytest.mark.skipif(PANDOC is None, reason="pandoc 不可用")
@pytest.mark.parametrize(
    "tex",
    [
        r'$$     H _ { 0 } : { \cal D } = 0 .\tag{5}     $$',
        r'$Z = { \cal Y } .$',
        r'$$ C = \left( x \right) ^ { 1 / \mathnormal { p } } .\tag{8} $$',
        r'$$ R = \cfrac { 1 } { N } Y Y ^ { T } .\tag{10} $$',
        r'$$ e = \overline { { { \bf Y } } } .\tag{15} $$',
        # {\cal X}/{\bf X} 位于外层命令参数内：改写必须保留外层分组括号，
        # 否则 \tilde{\bf X} 变成 \tilde\mathbf{X}，texmath 解析失败、整条公式降级为文本
        r'$$ s = \sqrt{\cal X} $$',
        r'$$ t = \tilde{\bf X} - \tilde{\bf y} $$',
        r'$$ u = { \tilde { \bf X } } _ { 3 } $$',
        r'$$ b = \boldsymbol{\cal X} $$',
        r'$$ l = \left\{\cal X\right\} $$',
        # 声明 + 分组 / 命令参数形态
        r'$$ c = \cal{X} $$',
        r'$$ d = \bf{X} $$',
        r'$$ a = \cal\alpha $$',
        r'$$ g = {\bf\alpha} $$',
        r'$$ h = {\bf { \cal X }} $$',
    ],
)
def test_legacy_latex_commands_convert_cleanly(tex):
    # texmath 不支持上述旧式命令时，pandoc 会把整个公式按原始 TeX 文本渲染
    md = f"正文\n\n{tex}\n\n结尾\n"
    result = subprocess.run(
        [
            PANDOC,
            "-f",
            "markdown+tex_math_dollars+raw_tex+tex_math_double_backslash+tex_math_single_backslash",
            "-t", "docx", "-o", "-",
            "--lua-filter", str(LUA_FILTER),
        ],
        input=md.encode("utf-8"),
        capture_output=True,
    )
    stderr = result.stderr.decode("utf-8", "ignore")
    assert "Could not convert" not in stderr
