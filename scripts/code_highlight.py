"""把 Report Code 段落换成带底纹的 1×1 表格，并按 pygments 分词上色。

与旧脚本的区别：不再猜测文件是否被 Word 占用、也不再静默改名输出，
调用方明确指定输出路径；空行比较改成忽略空白，不再误报。
"""
from __future__ import annotations

import re
from pathlib import Path

from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from pygments import lex
from pygments.lexers import get_lexer_by_name
from pygments.token import Token
from pygments.util import ClassNotFound

import report_style as rs

FENCE_PATTERN = re.compile(r'^([ \t]*)```([^\n]*)\n(.*?)^[ \t]*```', re.M | re.S)

LEXER_ALIASES = {'sh': 'bash', 'shell': 'bash', 'text': 'text', 'plaintext': 'text'}

# 命令行工具名在 bash 词法里不算关键字，单独染成关键字色。
SHELL_COMMANDS = {'mvn', 'npm', 'node', 'mysql', 'cd', 'java', 'git', 'docker'}


def _dedent(body, indent):
    """按开启围栏的缩进量整体去缩进。

    要和 report_builder 写进文档的代码段落保持一致：那边按同样的缩进量去掉前缀，
    这里若不去，分词结果的每行会多出前导空格，与段落文本对不上。
    """
    width = len(indent)
    if not width:
        return body
    lines = []
    for line in body.split('\n'):
        lines.append(line[width:] if line[:width].strip() == '' else line)
    return '\n'.join(lines)


def fence_blocks(markdown_text):
    """取出 Markdown 里所有代码围栏，返回 [(语言, 内容)]。

    围栏允许带缩进——写在列表项里的代码块，围栏与内容一起缩进；内容按开启围栏的
    缩进量去缩进后再返回，与 report_builder 写进文档的代码段落逐行对应。
    """
    return [(language.strip(), _dedent(body, indent))
            for indent, language, body in FENCE_PATTERN.findall(markdown_text)]


def collect_code_paragraphs(document):
    """把正文里连续的 Report Code 段落归成块。"""
    blocks = []
    current = []
    for paragraph in document.paragraphs:
        if paragraph.style.name == 'Report Code':
            current.append(paragraph)
        elif current:
            blocks.append(current)
            current = []
    if current:
        blocks.append(current)
    return blocks


def token_kind(token_type, value, language):
    if token_type in Token.Comment:
        return 'comment', False
    if token_type in Token.Keyword or token_type in Token.Name.Builtin:
        return 'keyword', True
    if token_type in Token.Name.Decorator:
        return 'decorator', False
    if token_type in Token.Name.Function:
        return 'function', False
    if token_type in Token.Literal.String:
        return 'string', False
    if token_type in Token.Literal.Number:
        return 'number', False
    if token_type in Token.Operator:
        return 'operator', False
    if token_type in Token.Punctuation:
        return 'punctuation', False
    # diff 的增删行、markdown 的标题与强调、控制台提示符都归在 Generic 下。
    if token_type in Token.Generic.Inserted:
        return 'inserted', False
    if token_type in Token.Generic.Deleted or token_type in Token.Generic.Error:
        return 'deleted', False
    if (token_type in Token.Generic.Heading or token_type in Token.Generic.Subheading
            or token_type in Token.Generic.Strong or token_type in Token.Generic.Prompt):
        return 'keyword', True
    if token_type in Token.Generic.Emph:
        return 'text', False
    if language == 'sh' and value.strip() in SHELL_COMMANDS:
        return 'keyword', True
    return 'text', False


def styled_run(paragraph, text, kind, bold, config):
    run = paragraph.add_run(text)
    run.font.size = Pt(config['sizes']['code'])
    run.bold = bold
    mono = config['code']['font'] == 'mono'
    rs.set_run_fonts(run._element.get_or_add_rPr(), config, bold, mono)
    color = run._element.get_or_add_rPr().find(qn('w:color'))
    if color is not None:
        color.set(qn('w:val'), config['code']['palette'][kind])
    return run


def _plain_lines(source):
    return [[(line, 'text', False)] if line else [] for line in source.split('\n')]


def _diff_lines(source):
    """diff 按行首符号上色。

    不能交给 pygments：它的 diff 词法器会按 ---/+++ 里的文件名后缀挑子词法器，
    内容不是那个语言时就全部退化成 Token.Error，增删行反而没有颜色。
    """
    lines = []
    for line in source.split('\n'):
        if line.startswith('+++') or line.startswith('---'):
            lines.append([(line, 'text', True)])
        elif line.startswith('@@'):
            lines.append([(line, 'keyword', True)])
        elif line.startswith('+'):
            lines.append([(line, 'inserted', False)])
        elif line.startswith('-'):
            lines.append([(line, 'deleted', False)])
        else:
            lines.append([(line, 'text', False)] if line else [])
    return lines


def _console_lines(source):
    """控制台只把提示符染成关键字色。

    pygments 的 console 词法器会吞掉提示符之后的内容，不能用。
    """
    lines = []
    for line in source.split('\n'):
        for prompt in ('$ ', '> ', '# '):
            if line.startswith(prompt):
                rest = line[len(prompt):]
                tokens = [(prompt, 'keyword', True)]
                if rest:
                    tokens.append((rest, 'text', False))
                lines.append(tokens)
                break
        else:
            lines.append([(line, 'text', False)] if line else [])
    return lines


def lex_lines(code, language, config):
    """整块分词后按行切开，返回 (每行的 [(文本, 类别, 加粗)], 是否退回纯文本)。

    整块一起分词，注释和字符串的跨行状态才不会丢；再切回一行一段，
    这样 Report Code 的悬挂缩进只作用于折行，不会把每一行都缩进去。
    """
    source = '\n'.join(code.splitlines())
    if language == 'diff':
        return _diff_lines(source), False
    if language == 'console':
        return _console_lines(source), False
    name = LEXER_ALIASES.get(language, language or 'text')
    try:
        lexer = get_lexer_by_name(name, stripnl=False, ensurenl=False)
    except ClassNotFound:
        return _plain_lines(source), True
    lines = [[]]
    pieces = []
    for token_type, value in lex(source, lexer):
        if not value:
            continue
        pieces.append(value)
        kind, bold = token_kind(token_type, value, language)
        for position, part in enumerate(value.split('\n')):
            if position:
                lines.append([])
            if part:
                lines[-1].append((part, kind, bold))
    # 少数词法器（例如 console）会吞掉内容，宁可不着色也不能丢字。
    if ''.join(pieces) != source:
        return _plain_lines(source), True
    return lines, False


def build_frame(document, config):
    """建一个 1×1 表格作为代码框。"""
    code = config['code']
    table = document.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    props = table._tbl.tblPr

    width = OxmlElement('w:tblW')
    width.set(qn('w:w'), str(round(rs.content_width_cm(config) / 2.54 * 1440)))
    width.set(qn('w:type'), 'dxa')
    props.append(width)

    borders = OxmlElement('w:tblBorders')
    for name in ('top', 'left', 'bottom', 'right'):
        border = OxmlElement('w:' + name)
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '6')
        border.set(qn('w:color'), code['border_color'])
        borders.append(border)
    props.append(borders)

    margins = OxmlElement('w:tblCellMar')
    for side, value in (('top', 135), ('bottom', 135), ('left', 165), ('right', 165)):
        item = OxmlElement('w:' + side)
        item.set(qn('w:w'), str(value))
        item.set(qn('w:type'), 'dxa')
        margins.append(item)
    props.append(margins)
    return table


def frame_code_blocks(document, markdown_path, config):
    """入口：把代码段落装箱上色，返回处理结果摘要。"""
    if not config['code']['frame']:
        return {'framed': 0, 'skipped': 'code.frame 为 false'}

    source = Path(markdown_path).read_text(encoding='utf-8')
    fences = fence_blocks(source)
    blocks = collect_code_paragraphs(document)
    if len(blocks) != len(fences):
        raise RuntimeError(
            f'代码块数量对不上：Word 里 {len(blocks)} 个，Markdown 里 {len(fences)} 个。'
            '请检查 Markdown 的 ``` 是否成对。'
        )

    mismatched = []
    for index, (block, (language, body)) in enumerate(zip(blocks, fences), 1):
        expected = ''.join(body.split())
        actual = ''.join(''.join(p.text for p in block).split())
        if expected != actual:
            mismatched.append(index)

    plain_fallback = 0
    for block, (language, body) in zip(blocks, fences):
        lines, fell_back = lex_lines(body, language, config)
        if fell_back:
            plain_fallback += 1
        if len(lines) != len(block):
            raise RuntimeError(
                f'代码块按行切开得到 {len(lines)} 行，但 Word 里有 {len(block)} 段，无法逐行上色。'
            )
        for paragraph, tokens in zip(block, lines):
            for run in list(paragraph.runs):
                run._element.getparent().remove(run._element)
            for text, kind, bold in tokens:
                styled_run(paragraph, text, kind, bold, config)
            paragraph.paragraph_format.line_spacing = Pt(config['sizes']['code'] * 1.22)

        table = build_frame(document, config)
        row = table.rows[0]
        row_props = row._tr.get_or_add_trPr()
        cant_split = OxmlElement('w:cantSplit')
        if len(block) > config['code']['split_after_lines']:
            cant_split.set(qn('w:val'), '0')
        row_props.append(cant_split)

        cell = table.cell(0, 0)
        for existing in list(cell.paragraphs):
            existing._p.getparent().remove(existing._p)
        shading = OxmlElement('w:shd')
        shading.set(qn('w:fill'), config['code']['shading'])
        cell._tc.get_or_add_tcPr().append(shading)

        block[0]._p.addprevious(table._tbl)
        for paragraph in block:
            cell._tc.append(paragraph._p)

    return {'framed': len(blocks), 'mismatched': mismatched, 'plain_fallback': plain_fallback}
