"""把 Markdown 报告转换成 Word 文档。

支持的 Markdown 语法见 references/config.md；不支持的语法不会报错，
会按普通正文处理。转换结果与旧脚本逐项对齐（同样的边框、边距、字体、
分页规则），另外补上了无序列表、引用、分隔线、斜体、删除线和行内代码等宽字体。
"""
from __future__ import annotations

import re
from pathlib import Path
from urllib.parse import unquote

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.shared import Cm, Pt
from PIL import Image

import report_style as rs

# 行内标记，长的放前面，否则 ** 会被 * 先吃掉。
INLINE_PATTERN = re.compile(
    r'(\*\*\*[^*]+\*\*\*'
    r'|\*\*[^*]+\*\*'
    r'|~~[^~]+~~'
    r'|`[^`]+`'
    r'|\[[^\]]+\]\([^)]+\)'
    r'|\*[^*\n]+\*'
    r'|(?<![\w`])_[^_\n]+_(?![\w`]))'
)

HEADING_PATTERN = re.compile(r'^(#{2,4}) (.+)$')
IMAGE_PATTERN = re.compile(r'^!\[([^\]]*)\]\(([^)]+)\)$')
BULLET_PATTERN = re.compile(r'^(\s*)[-*+] +(.*)$')
RULE_PATTERN = re.compile(r'^(-{3,}|\*{3,}|_{3,})$')
SEPARATOR_PATTERN = re.compile(r'^[\s:\-|]+$')


def split_row(line):
    """拆一行表格，处理转义的竖线。"""
    text = line.strip()
    if text.startswith('|'):
        text = text[1:]
    if text.endswith('|') and not text.endswith('\\|'):
        text = text[:-1]
    return [cell.replace('\\|', '|').strip() for cell in re.split(r'(?<!\\)\|', text)]


def is_separator_row(line):
    text = line.strip()
    return '-' in text and bool(SEPARATOR_PATTERN.match(text))


class ReportBuilder:
    def __init__(self, config, source_path):
        self.config = config
        self.source_path = Path(source_path).resolve()
        self.source_dir = self.source_path.parent
        captions = config['captions']
        # patterns 是旧配置里的通用列表，图和表都并入，保证老配置照旧可用。
        legacy = captions['patterns']
        self.figure_captions = [re.compile(p) for p in captions['figure_patterns'] + legacy]
        self.table_captions = [re.compile(p) for p in captions['table_patterns'] + legacy]
        self.reference_pattern = re.compile(config['references']['pattern'])
        self.number_pattern = re.compile(config['lists']['number_pattern'])
        self.code_languages = []
        self.warnings = []

        self.document = Document()
        self.section = rs.apply_page(self.document, config)
        rs.define_styles(self.document, config)
        rs.enable_field_updates(self.document)
        rs.add_page_number_footer(self.section, config)
        rs.apply_metadata(self.document, config)

    # ------------------------------------------------------------------
    # 行内标记
    # ------------------------------------------------------------------

    def add_run(self, paragraph, text, bold=False, italic=False, strike=False,
                mono=False, size=None):
        run = paragraph.add_run(text)
        run.bold = bold
        run.italic = italic
        if strike:
            run.font.strike = True
        rs.set_run_fonts(run._element.get_or_add_rPr(), self.config, bold, mono)
        if size:
            run.font.size = Pt(size)
        elif mono:
            run.font.size = Pt(self.config['sizes']['code'])
        return run

    def add_hyperlink(self, paragraph, label, target, bold=False, size=None):
        hyperlink = OxmlElement('w:hyperlink')
        hyperlink.set(qn('r:id'), paragraph.part.relate_to(target, RT.HYPERLINK, is_external=True))
        element = OxmlElement('w:r')
        rpr = OxmlElement('w:rPr')
        rs.set_run_fonts(rpr, self.config, bold)
        if bold:
            rpr.append(OxmlElement('w:b'))
        if size:
            width = OxmlElement('w:sz')
            width.set(qn('w:val'), str(int(size * 2)))
            rpr.append(width)
        element.append(rpr)
        text = OxmlElement('w:t')
        text.text = label
        element.append(text)
        hyperlink.append(element)
        paragraph._p.append(hyperlink)

    def add_inline(self, paragraph, text, bold=False, size=None):
        for part in INLINE_PATTERN.split(text or ''):
            if not part:
                continue
            is_bold, is_italic, is_strike, is_mono = bold, False, False, False
            if part.startswith('***') and part.endswith('***') and len(part) > 6:
                part, is_bold, is_italic = part[3:-3], True, True
            elif part.startswith('**') and part.endswith('**') and len(part) > 4:
                part, is_bold = part[2:-2], True
            elif part.startswith('~~') and part.endswith('~~') and len(part) > 4:
                part, is_strike = part[2:-2], True
            elif part.startswith('`') and part.endswith('`') and len(part) > 2:
                part, is_mono = part[1:-1], True
            elif part.startswith('*') and part.endswith('*') and len(part) > 2:
                part, is_italic = part[1:-1], True
            elif part.startswith('_') and part.endswith('_') and len(part) > 2:
                part, is_italic = part[1:-1], True

            link = re.fullmatch(r'\[([^\]]+)\]\(([^)]+)\)', part)
            if link:
                self.add_hyperlink(paragraph, link.group(1), link.group(2), is_bold, size)
                continue
            self.add_run(paragraph, part, is_bold, is_italic, is_strike, is_mono, size)

    # ------------------------------------------------------------------
    # 段落级结构
    # ------------------------------------------------------------------

    def flush(self, buffer):
        if not buffer:
            return
        paragraph = self.document.add_paragraph()
        self.add_inline(paragraph, ' '.join(buffer))
        buffer.clear()

    def previous_has_drawing(self):
        for paragraph in reversed(self.document.paragraphs):
            if paragraph.text.strip():
                return bool(paragraph._p.findall('.//' + qn('w:drawing')))
            if paragraph._p.findall('.//' + qn('w:drawing')):
                return True
        return False

    def next_starts_table(self, lines, index):
        """这一行之后（跳过空行）的第一个非空行是不是表格。"""
        following = index + 1
        while following < len(lines) and not lines[following].strip():
            following += 1
        return following < len(lines) and lines[following].strip().startswith('|')

    def caption_kind(self, lines, index):
        """判断这一行是不是题注，返回 'figure' / 'table' / None。

        图题注写在图片下面，表题注写在表格上面。题注必须贴着图或表——正文里
        出现的「图 3-1」「表 3-1」正是靠这一条才不会被误判成题注。
        """
        text = lines[index].strip()
        captions = self.config['captions']
        after_image = self.previous_has_drawing()
        before_table = self.next_starts_table(lines, index)
        if (any(pattern.match(text) for pattern in self.figure_captions)
                and (after_image or not captions['require_preceding_image'])):
            return 'figure'
        if (any(pattern.match(text) for pattern in self.table_captions)
                and (before_table or not captions['require_following_table'])):
            return 'table'
        return None

    def add_cover(self, lines, index):
        text = lines[index][2:].strip()
        marker = self.config['cover']['break_before']
        if marker and marker in text:
            text = text.replace(marker, '\n' + marker, 1)
        paragraph = self.document.add_paragraph(style='Title')
        self.add_inline(paragraph, text, True)
        index += 1

        if self.config['cover']['subtitle_from_next_line']:
            while index < len(lines) and not lines[index].strip():
                index += 1
            if index < len(lines):
                subtitle = self.document.add_paragraph(style='Subtitle')
                self.add_inline(subtitle, lines[index].strip())
                index += 1

        toc_title = self.config['cover']['toc_title']
        if toc_title:
            heading = rs.blank_paragraph(self.document, self.config, space_after=8)
            heading.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            self.add_inline(heading, toc_title, True, self.config['sizes']['h1'])
            self.add_toc_field()
        return index

    def add_toc_field(self):
        paragraph = rs.blank_paragraph(self.document, self.config)
        run = paragraph.add_run()
        rs.set_run_fonts(run._element.get_or_add_rPr(), self.config)
        begin = OxmlElement('w:fldChar')
        begin.set(qn('w:fldCharType'), 'begin')
        instruction = OxmlElement('w:instrText')
        instruction.set(qn('xml:space'), 'preserve')
        instruction.text = ' TOC \\o "1-2" \\h \\z \\u '
        separate = OxmlElement('w:fldChar')
        separate.set(qn('w:fldCharType'), 'separate')
        placeholder = OxmlElement('w:t')
        placeholder.text = self.config['cover']['toc_title'] or '目录'
        end = OxmlElement('w:fldChar')
        end.set(qn('w:fldCharType'), 'end')
        run._r.extend([begin, instruction, separate, placeholder, end])

    def add_heading(self, level, text, first_body):
        paragraph = self.document.add_paragraph(style=f'Heading {level}')
        headings = self.config['headings']
        break_here = False
        if first_body and headings['page_break_level1']:
            break_here = True
        elif level == 1 and any(text.startswith(prefix) for prefix in headings['page_break_prefixes']):
            break_here = True
        if break_here:
            paragraph.paragraph_format.page_break_before = True
        self.add_inline(paragraph, text, True)

    def add_code_block(self, lines, index):
        # 围栏可以带缩进：写在有序/无序列表项里的代码块，围栏和内容一起缩进若干个空格。
        # 结束围栏同样带缩进，所以按 strip 后的结果判断围栏；内容按开启围栏的缩进量
        # 整体去缩进，否则每一行都会多出一截空白。
        fence = lines[index]
        indent = len(fence) - len(fence.lstrip())
        language = fence.strip()[3:].strip()
        self.code_languages.append(language)
        index += 1
        while index < len(lines) and not lines[index].strip().startswith('```'):
            text = lines[index]
            if indent and text[:indent].strip() == '':
                text = text[indent:]
            paragraph = self.document.add_paragraph(style='Report Code')
            run = paragraph.add_run(text or ' ')
            rs.set_run_fonts(run._element.get_or_add_rPr(), self.config,
                             mono=self.config['code']['font'] == 'mono')
            run.font.size = Pt(self.config['sizes']['code'])
            paragraph.paragraph_format.keep_with_next = False
            if (self.config['code']['keep_with_next_for_text'] and language == 'text'
                    and index + 1 < len(lines) and lines[index + 1].startswith(' ')):
                paragraph.paragraph_format.keep_with_next = True
            index += 1
        index += 1
        gap = rs.blank_paragraph(self.document, self.config, line_spacing=Pt(4))
        gap.paragraph_format.space_after = Pt(2)
        gap.add_run().font.size = Pt(4)
        return index

    def add_table(self, rows):
        config = self.config
        tables = config['tables']
        widths = rs.table_widths(rows, config)
        columns = len(rows[0])

        table = self.document.add_table(rows=0, cols=columns)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        props = table._tbl.tblPr

        borders = OxmlElement('w:tblBorders')
        for name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            border = OxmlElement('w:' + name)
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '5')
            border.set(qn('w:color'), tables['border_color'])
            borders.append(border)
        props.append(borders)

        margins = OxmlElement('w:tblCellMar')
        for side, value in (('top', 65), ('bottom', 65), ('left', 85), ('right', 85)):
            item = OxmlElement('w:' + side)
            item.set(qn('w:w'), str(value))
            item.set(qn('w:type'), 'dxa')
            margins.append(item)
        props.append(margins)

        for column, width in zip(table.columns, widths):
            column.width = Cm(width)

        center_values = set(tables['center_values'])
        for row_index, values in enumerate(rows):
            row = table.add_row()
            row_props = row._tr.get_or_add_trPr()
            row_props.append(OxmlElement('w:cantSplit'))
            if row_index == 0:
                row_props.append(OxmlElement('w:tblHeader'))
            for index, value in enumerate(values):
                cell = row.cells[index]
                cell.width = Cm(widths[index])
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                if row_index == 0:
                    shading = OxmlElement('w:shd')
                    shading.set(qn('w:fill'), tables['header_fill'])
                    cell._tc.get_or_add_tcPr().append(shading)
                paragraph = cell.paragraphs[0]
                paragraph.style = 'Report Table'
                paragraph.paragraph_format.keep_with_next = row_index < 2
                paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                if (row_index == 0 or value in center_values
                        or re.fullmatch(r'[\d,.]+', value)):
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                self.add_inline(paragraph, value, row_index == 0, config['sizes']['table'])

        gap = rs.blank_paragraph(self.document, self.config, line_spacing=Pt(4))
        gap.paragraph_format.space_after = Pt(0)
        gap.add_run().font.size = Pt(4)

    def resolve_image(self, raw):
        candidate = unquote(raw.strip().strip('<>'))
        options = [self.source_dir / candidate, Path(candidate), self.source_dir / Path(candidate).name]
        for option in options:
            if option.is_file():
                return option
        return None

    def add_image(self, raw, has_caption):
        path = self.resolve_image(raw)
        if path is None:
            self.warnings.append(f'找不到图片：{raw}')
            paragraph = self.document.add_paragraph()
            self.add_inline(paragraph, f'[缺失图片：{raw}]')
            return
        with Image.open(path) as image:
            ratio = image.height / image.width
        images = self.config['images']
        max_width = images['max_width_cm'] or rs.content_width_cm(self.config)
        width = min(max_width, images['max_height_cm'] / ratio)
        paragraph = rs.blank_paragraph(self.document, self.config, space_before=6)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        paragraph.paragraph_format.keep_with_next = bool(has_caption)
        paragraph.paragraph_format.line_spacing = 1
        paragraph.add_run().add_picture(str(path), width=Cm(width))

    def add_caption(self, text, keep_with_next=False):
        # 表题注写在表格上面，要跟着下面的表格走，否则容易被留在上一页页脚处。
        paragraph = self.document.add_paragraph(style='Caption')
        paragraph.paragraph_format.keep_with_next = bool(keep_with_next)
        self.add_inline(paragraph, text)

    def add_reference(self, text):
        paragraph = self.document.add_paragraph()
        fmt = paragraph.paragraph_format
        fmt.first_line_indent = Pt(0)
        indent = self.config['references']['hanging_indent_cm']
        fmt.left_indent = Cm(indent)
        fmt.first_line_indent = Cm(-indent)
        fmt.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        self.add_inline(paragraph, text)

    def add_numbered(self, text):
        paragraph = self.document.add_paragraph()
        fmt = paragraph.paragraph_format
        indent = self.config['lists']['number_hanging_indent_cm']
        fmt.left_indent = Cm(indent)
        fmt.first_line_indent = Cm(-indent)
        fmt.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        self.add_inline(paragraph, text)

    def bullet_level(self, line):
        match = BULLET_PATTERN.match(line)
        if not match:
            return None
        indent = len(match.group(1).replace('\t', '    '))
        level = 0 if indent < 2 else (1 if indent < 4 else 2)
        chars = self.config['lists']['bullet_chars']
        return min(level, len(chars) - 1), match.group(2)

    def add_bullet(self, text, level):
        config = self.config
        indent = config['lists']['bullet_indent_cm']
        paragraph = self.document.add_paragraph()
        fmt = paragraph.paragraph_format
        fmt.left_indent = Cm(indent * (level + 1))
        fmt.first_line_indent = Cm(-indent)
        fmt.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        fmt.tab_stops.add_tab_stop(Cm(indent * (level + 1)), WD_TAB_ALIGNMENT.LEFT)
        marker = config['lists']['bullet_chars'][level]
        self.add_inline(paragraph, f'{marker}\t{text}')

    def add_quote(self, text):
        paragraph = self.document.add_paragraph()
        fmt = paragraph.paragraph_format
        indent = self.config['lists']['quote_indent_cm']
        fmt.left_indent = Cm(indent)
        fmt.first_line_indent = Pt(0)
        self.paragraph_border(paragraph, 'left', 'BFBFBF', 12)
        self.add_inline(paragraph, text)

    def add_rule(self):
        paragraph = rs.blank_paragraph(self.document, self.config, space_before=6, space_after=6)
        self.paragraph_border(paragraph, 'bottom', 'BFBFBF', 6)

    def paragraph_border(self, paragraph, edge, color, size):
        ppr = paragraph._p.get_or_add_pPr()
        borders = ppr.find(qn('w:pBdr'))
        if borders is None:
            borders = OxmlElement('w:pBdr')
            ppr.append(borders)
        border = OxmlElement('w:' + edge)
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), str(size))
        border.set(qn('w:space'), '4')
        border.set(qn('w:color'), color)
        borders.append(border)

    # ------------------------------------------------------------------
    # 主循环
    # ------------------------------------------------------------------

    def build(self, lines):
        config = self.config
        index = 0
        first_body = True
        buffer = []
        while index < len(lines):
            line = lines[index]
            stripped = line.strip()
            if not stripped:
                self.flush(buffer)
                index += 1
                continue
            if stripped.startswith('```'):
                self.flush(buffer)
                index = self.add_code_block(lines, index)
                continue
            if stripped.startswith('# '):
                self.flush(buffer)
                index = self.add_cover(lines, index)
                continue
            heading = HEADING_PATTERN.match(stripped)
            if heading:
                self.flush(buffer)
                self.add_heading(len(heading.group(1)) - 1, heading.group(2), first_body)
                first_body = False
                index += 1
                continue
            if stripped.startswith('|'):
                self.flush(buffer)
                index = self.add_table_block(lines, index)
                continue
            image = IMAGE_PATTERN.match(stripped)
            if image:
                self.flush(buffer)
                following = index + 1
                while following < len(lines) and not lines[following].strip():
                    following += 1
                has_caption = (following < len(lines)
                               and any(p.match(lines[following].strip())
                                       for p in self.figure_captions))
                self.add_image(image.group(2), has_caption)
                index += 1
                continue
            if RULE_PATTERN.match(stripped):
                self.flush(buffer)
                self.add_rule()
                index += 1
                continue
            bullet = self.bullet_level(line)
            if bullet:
                self.flush(buffer)
                self.add_bullet(bullet[1], bullet[0])
                index += 1
                continue
            if self.number_pattern.match(stripped):
                self.flush(buffer)
                self.add_numbered(stripped)
                index += 1
                continue
            if self.reference_pattern.match(stripped):
                self.flush(buffer)
                self.add_reference(stripped)
                index += 1
                continue
            if stripped.startswith('>'):
                self.flush(buffer)
                self.add_quote(stripped.lstrip('>').strip())
                index += 1
                continue
            kind = self.caption_kind(lines, index)
            if kind:
                self.flush(buffer)
                self.add_caption(stripped, keep_with_next=(kind == 'table'))
                index += 1
                continue
            buffer.append(stripped)
            index += 1
        self.flush(buffer)
        return self.document

    def add_table_block(self, lines, index):
        rows = []
        while index < len(lines) and lines[index].strip().startswith('|'):
            stripped = lines[index].strip()
            if is_separator_row(stripped):
                index += 1
                continue
            rows.append(split_row(stripped))
            index += 1
        if not rows:
            return index
        columns = max(len(row) for row in rows)
        if any(len(row) != columns for row in rows):
            self.warnings.append(f'表格列数不一致（最多 {columns} 列），已按最多列补齐。')
        normalized = [row + [''] * (columns - len(row)) for row in rows]
        self.add_table(normalized)
        return index


def build_document(markdown_path, config):
    source = Path(markdown_path).resolve()
    lines = source.read_text(encoding='utf-8').splitlines()
    builder = ReportBuilder(config, source)
    document = builder.build(lines)
    return document, builder
