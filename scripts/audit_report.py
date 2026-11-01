"""转换结果自检。

检查项分两类：
  硬错误  结构性问题，说明转换有 bug（文本丢失、表宽合计不对、代码块数量对不上）
  警告    需要人工确认但未必是错（颜色越界、PDF 里有文字溢出页面）

pdfplumber 可选：装了才做逐字符溢出检查，没装就跳过。
"""
from __future__ import annotations

import re
from pathlib import Path

from docx.oxml.ns import qn
from docx.shared import Cm
from docx.text.paragraph import Paragraph

import code_highlight
import report_style as rs

EMU_PER_CM = 360000


def all_paragraphs(document):
    """遍历正文和表格单元格里的全部段落。"""
    for node in document.element.body.iter(qn('w:p')):
        yield Paragraph(node, document)


def body_text(document):
    """正文文本，跳过目录域、代码段落和空白段。

    代码段落在装箱时会被合并成一段，段落划分必然变化，所以单独用
    check_code_blocks 逐块比对内容，这里只负责正文。
    """
    texts = []
    for paragraph in all_paragraphs(document):
        if paragraph.style.name.lower().startswith('toc'):
            continue
        if paragraph.style.name == 'Report Code':
            continue
        if paragraph.text.strip():
            texts.append(paragraph.text)
    return texts


def code_tables(document):
    return [
        table for table in document.tables
        if len(table.rows) == 1 and len(table.columns) == 1
        and any(p.style.name == 'Report Code' for p in table.cell(0, 0).paragraphs)
    ]


def grid_width_cm(table):
    """读 tblGrid 里的列宽合计（厘米）；没有 grid 就返回 None。"""
    grid = table._tbl.find(qn('w:tblGrid'))
    if grid is None:
        return None
    columns = grid.findall(qn('w:gridCol'))
    if not columns:
        return None
    total = 0
    for column in columns:
        total += int(column.get(qn('w:w')) or 0)
    return total / 1440 * 2.54


def check_structure(document, config):
    errors, warnings = [], []
    frames = code_tables(document)
    data_tables = [t for t in document.tables if t not in frames]
    target = rs.content_width_cm(config)

    for index, table in enumerate(data_tables, 1):
        width = grid_width_cm(table)
        if width is None:
            warnings.append(f'第 {index} 张数据表没有 tblGrid，无法核对列宽。')
        elif abs(width - target) > 0.05:
            errors.append(f'第 {index} 张数据表列宽合计 {width:.2f}cm，应为 {target:.2f}cm。')

    for index, shape in enumerate(document.inline_shapes, 1):
        width = shape.width / EMU_PER_CM
        if width > target + 0.05:
            errors.append(f'第 {index} 张图片宽 {width:.2f}cm，超出正文宽度 {target:.2f}cm。')

    return {
        'errors': errors,
        'warnings': warnings,
        'dataTables': len(data_tables),
        'codeTables': len(frames),
        'images': len(document.inline_shapes),
    }


def check_colors(document, config):
    """代码段落只能用配色表里的颜色，其余段落必须是黑色。"""
    allowed = {'000000', 'auto'} | set(config['code']['palette'].values())
    errors = []
    counts = {}
    for paragraph in all_paragraphs(document):
        is_code = paragraph.style.name == 'Report Code'
        for run in paragraph.runs:
            if not run.text:
                continue
            node = run._r.get_or_add_rPr().find(qn('w:color'))
            color = (node.get(qn('w:val')) if node is not None else '000000') or '000000'
            counts[color] = counts.get(color, 0) + 1
            if is_code and color not in allowed:
                errors.append(f'代码段落出现配色表外的颜色 {color}：{run.text[:30]!r}')
            elif not is_code and color not in {'000000', 'auto'}:
                errors.append(f'正文段落被染成 {color}：{run.text[:30]!r}')
    return {'errors': errors, 'colors': counts}


def check_code_blocks(document, markdown_path, config):
    """代码块数量与内容必须和 Markdown 围栏一致。

    装箱后代码段落已经移进单元格，按代码框表格取块最可靠。
    """
    fences = code_highlight.fence_blocks(Path(markdown_path).read_text(encoding='utf-8'))
    tables = code_tables(document)
    errors = []
    if len(tables) != len(fences):
        errors.append(f'代码块数量不一致：Word {len(tables)} 个，Markdown {len(fences)} 个。')
        return {'errors': errors, 'count': len(tables)}
    mismatched = []
    for index, (table, (language, body)) in enumerate(zip(tables, fences), 1):
        actual = ''.join(paragraph.text for paragraph in table.cell(0, 0).paragraphs)
        if ''.join(actual.split()) != ''.join(body.split()):
            mismatched.append(index)
    if mismatched:
        errors.append(f'第 {mismatched} 个代码块内容与 Markdown 不一致。')
    return {'errors': errors, 'count': len(tables), 'mismatched': mismatched}


def check_text_preserved(before, after):
    """加代码框前后，正文文本必须完全一致。"""
    errors = []
    if before != after:
        missing = [line for line in before if line not in set(after)]
        added = [line for line in after if line not in set(before)]
        if missing:
            errors.append(f'有 {len(missing)} 行文本在加代码框后丢失，例如：{missing[0][:60]!r}')
        if added:
            errors.append(f'出现了 {len(added)} 行原本没有的文本，例如：{added[0][:60]!r}')
        if not missing and not added:
            errors.append('正文段落顺序或数量发生了变化。')
    return {'errors': errors, 'paragraphs': len(after)}


def check_pdf_overflow(pdf_path):
    """可选：逐字符检查有没有内容跑出页面。需要 pdfplumber。"""
    try:
        import pdfplumber
    except ImportError:
        return {'errors': [], 'skipped': 'pdfplumber 未安装，跳过 PDF 溢出检查'}
    overflow = []
    pages = 0
    with pdfplumber.open(pdf_path) as pdf:
        for number, page in enumerate(pdf.pages, 1):
            pages = number
            for char in page.chars:
                if (char['x0'] < -1 or char['x1'] > page.width + 1
                        or char['top'] < -1 or char['bottom'] > page.height + 1):
                    overflow.append((number, char['text']))
    errors = []
    if overflow:
        sample = ', '.join(f'第{p}页 {c!r}' for p, c in overflow[:5])
        errors.append(f'有 {len(overflow)} 个字符跑出页面：{sample}')
    return {'errors': errors, 'pages': pages}


def audit(document, before_text, markdown_path, config, pdf_path=None):
    """跑完全部检查，返回报告字典。"""
    after_text = body_text(document)
    report = {
        'structure': check_structure(document, config),
        'colors': check_colors(document, config),
        'code': check_code_blocks(document, markdown_path, config),
        'text': check_text_preserved(before_text, after_text),
    }
    if pdf_path:
        report['pdf'] = check_pdf_overflow(pdf_path)

    errors = []
    warnings = []
    for name, section in report.items():
        errors.extend(f'[{name}] {message}' for message in section.get('errors', []))
        warnings.extend(f'[{name}] {message}' for message in section.get('warnings', []))
    report['errors'] = errors
    report['warnings'] = warnings
    return report
