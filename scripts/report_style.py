"""配置加载与排版样式定义。

内置默认值就是中文课程报告的排版：A4、2.5cm 边距、宋体正文、黑体加粗、
Times New Roman 西文、首行缩进 2 字符、两端对齐。

在 Markdown 旁边放一个同名的 .toml / .json / .yaml 就能覆盖任意一项，
只写要改的部分即可，其余继续用默认值。
"""
from __future__ import annotations

import copy
import json
import tomllib
import unicodedata
from pathlib import Path

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

DEFAULTS = {
    'page': {
        'width_cm': 21.0,
        'height_cm': 29.7,
        'margin_cm': 2.5,
        'header_distance_cm': 1.2,
        'footer_distance_cm': 1.3,
    },
    'fonts': {
        'latin': 'Times New Roman',
        'cjk': '宋体',
        'cjk_bold': '黑体',
        'mono': 'Consolas',
    },
    'sizes': {
        'title': 20.0,
        'subtitle': 12.0,
        'body': 12.0,
        'h1': 16.0,
        'h2': 14.0,
        'h3': 12.0,
        'caption': 10.5,
        'code': 10.5,
        'table': 12.0,
        'toc': 12.0,
        'footer': 10.5,
    },
    'body': {
        'first_line_indent_pt': 24.0,
        'line_spacing': 1.5,
        'align': 'justify',
    },
    'headings': {
        'line_spacing': 1.35,
        'space_before': [14.0, 10.0, 10.0],
        'space_after': 6.0,
        # 正文第一个标题另起一页，让封面和目录单独占一页。
        'page_break_level1': True,
        # 一级标题以这些词开头时也另起一页。
        'page_break_prefixes': ['附录 '],
    },
    'cover': {
        # 标题里出现这个词时前面换行；留空字符串表示不换行。
        'break_before': '课程设计报告',
        'toc_title': '目录',
        # True 表示把 # 标题之后的第一个非空行当副标题。
        'subtitle_from_next_line': True,
    },
    'captions': {
        'patterns': ['^图 \\d+-\\d+ ', '^表 \\d+-\\d+ '],
        # True 表示题注只在紧跟在图片之后时才生效，避免正文里的“图 3-1”被误判。
        'require_preceding_image': True,
    },
    'references': {
        'pattern': '^\\[\\d+\\] ',
        'hanging_indent_cm': 0.65,
    },
    'images': {
        # 0 表示用正文宽度（页宽减两倍边距）。
        'max_width_cm': 0.0,
        # 图片再高也不会超过这个高度，宽度按高宽比反推。
        'max_height_cm': 16.0,
    },
    'lists': {
        'bullet_chars': ['•', '◦', '▪'],
        'bullet_indent_cm': 0.74,
        'number_pattern': '^\\d+\\. ',
        'number_hanging_indent_cm': 0.6,
        'quote_indent_cm': 0.74,
    },
    'tables': {
        # auto 按内容自动分配；equal 均分；rules 用下面的表头规则。
        'width_mode': 'auto',
        'minimum_column_cm': 1.4,
        # 单个单元格再长也不会把该列撑过这个权重（按半角字符数计）。
        'maximum_weight': 40.0,
        'header_fill': 'F2F2F2',
        'border_color': 'D9D9D9',
        # rules 模式：键是表头用逗号连接，值是各列宽度（厘米）。
        'rules': {},
        'center_values': ['是', '否', '1:1', '1:N', 'M:N', 'ref', 'index'],
    },
    'code': {
        'frame': True,
        'shading': 'F6F8FA',
        'border_color': 'D9D9D9',
        # 超过这个行数就允许表格跨页断开，否则整块保持在同一页。
        'split_after_lines': 30,
        # text 块里对齐的 ASCII 内容，下一行以空格开头时不要断开。
        'keep_with_next_for_text': True,
        # mono 用 fonts.mono（等宽，代码更易读）；latin 沿用正文字体。
        # 无论哪种，代码里的中文都用 fonts.cjk，因为等宽字体没有汉字字形。
        'font': 'mono',
        # 代码行折行后的缩进量，让折下来的部分与上一行区分开。
        'wrap_indent_cm': 0.6,
        'palette': {
            'text': '24292F',
            'keyword': '1F4E79',
            'string': '985C10',
            'comment': '58724C',
            'number': '7C3F8C',
            'function': '006D77',
            'decorator': '6F42A0',
            'operator': '455265',
            'punctuation': '334155',
            # diff 的增删行、控制台提示符等 Generic 类记号用这两色。
            'inserted': '1A7F37',
            'deleted': 'CF222E',
        },
    },
    'metadata': {
        'title': '',
        'subject': '',
        'keywords': '',
    },
}

CONFIG_SUFFIXES = ('.toml', '.json', '.yaml', '.yml')

ALIGNMENTS = {
    'left': WD_ALIGN_PARAGRAPH.LEFT,
    'center': WD_ALIGN_PARAGRAPH.CENTER,
    'right': WD_ALIGN_PARAGRAPH.RIGHT,
    'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
}


class ConfigError(Exception):
    """配置读取失败，消息直接展示给使用者。"""


def deep_merge(base, override):
    """把 override 递归合并进 base 的副本，列表整体替换。"""
    result = copy.deepcopy(base)
    for key, value in (override or {}).items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = copy.deepcopy(value)
    return result


def read_config_file(path):
    suffix = path.suffix.lower()
    text = path.read_text(encoding='utf-8')
    if suffix == '.toml':
        return tomllib.loads(text)
    if suffix == '.json':
        return json.loads(text)
    if suffix in ('.yaml', '.yml'):
        try:
            import yaml
        except ImportError as error:
            raise ConfigError(
                f'{path.name} 是 YAML，但当前环境没有 PyYAML。'
                '请改成同名 .toml 或 .json 配置。'
            ) from error
        return yaml.safe_load(text) or {}
    raise ConfigError(f'不认识的配置后缀：{path.suffix}')


def find_config(markdown_path, explicit=None):
    """返回要加载的配置文件路径，找不到就返回 None。"""
    if explicit:
        candidate = Path(explicit).resolve()
        if not candidate.is_file():
            raise ConfigError(f'指定的配置文件不存在：{candidate}')
        return candidate
    for suffix in CONFIG_SUFFIXES:
        candidate = markdown_path.with_suffix(suffix)
        if candidate.is_file():
            return candidate
    return None


def load_config(markdown_path, explicit=None):
    """默认值 + 配置文件，返回 (配置字典, 实际用到的配置文件或 None)。"""
    path = find_config(markdown_path, explicit)
    if path is None:
        return copy.deepcopy(DEFAULTS), None
    override = read_config_file(path)
    if not isinstance(override, dict):
        raise ConfigError(f'{path.name} 的顶层必须是键值表。')
    return deep_merge(DEFAULTS, override), path


def content_width_cm(config):
    """正文可用宽度，表格和图片都按它排版。"""
    page = config['page']
    return round(page['width_cm'] - 2 * page['margin_cm'], 4)


# --------------------------------------------------------------------------
# 表格列宽
# --------------------------------------------------------------------------

def _is_wide(character):
    return unicodedata.east_asian_width(character) in ('W', 'F')


# 一个汉字正好是一个全角宽，记为 2。西文按 0.6 个汉字宽估算：Times New Roman
# 的大写字母和数字比 0.5em 宽，按 0.5 估会让 VARCHAR(64) 这类内容在列里折行。
ASCII_WEIGHT = 1.2
CJK_WEIGHT = 2.0


def measure(value):
    """把一个单元格的内容折算成半角字符宽度，一个汉字算 2。"""
    text = value or ''
    # Markdown 标记不占版面，先去掉再量。
    for token in ('**', '`', '~~'):
        text = text.replace(token, '')
    return sum(CJK_WEIGHT if _is_wide(character) else ASCII_WEIGHT for character in text)


def distribute(weights, total, minimum):
    """按权重分配 total，任何一列不低于 minimum，结果之和恒等于 total。"""
    count = len(weights)
    if count == 0:
        return []
    if minimum * count >= total:
        return [round(total / count, 4)] * count
    widths = [0.0] * count
    free = list(range(count))
    remaining = float(total)
    while free:
        pool = sum(weights[index] for index in free)
        scale = remaining / pool if pool > 0 else 0.0
        narrow = [index for index in free if weights[index] * scale < minimum]
        if not narrow:
            for index in free:
                widths[index] = weights[index] * scale
            break
        for index in narrow:
            widths[index] = minimum
            remaining -= minimum
            free.remove(index)
    # 浮点误差修到最后一列，保证合计正好是 total。
    widths[-1] += total - sum(widths)
    return [round(value, 4) for value in widths]


def auto_widths(rows, total, minimum, maximum_weight):
    """按每列最宽内容分配列宽。"""
    columns = len(rows[0])
    weights = []
    for index in range(columns):
        widest = max(measure(row[index]) for row in rows)
        weights.append(float(min(max(widest, 6), maximum_weight)))
    return distribute(weights, total, minimum)


def equal_widths(rows, total):
    columns = len(rows[0])
    return distribute([1.0] * columns, total, 0.0)


def rule_widths(rows, rules, total):
    """按配置里的表头规则取列宽，没有匹配的返回 None。"""
    header = ','.join(cell.strip() for cell in rows[0])
    for key, value in (rules or {}).items():
        if ','.join(part.strip() for part in key.split(',')) != header:
            continue
        widths = [float(item) for item in value]
        if len(widths) != len(rows[0]):
            raise ConfigError(
                f'表头规则「{key}」给了 {len(widths)} 个宽度，但表格有 {len(rows[0])} 列。'
            )
        return widths
    return None


def table_widths(rows, config):
    """算出这张表每列的宽度，永远返回与列数等长的列表。"""
    if not rows or not rows[0]:
        return []
    tables = config['tables']
    total = content_width_cm(config)
    columns = len(rows[0])

    def safe(widths):
        # 兜底：宽度数量对不上或合计偏差过大，就退回均分，绝不让上层崩掉。
        if not widths or len(widths) != columns:
            return equal_widths(rows, total)
        if abs(sum(widths) - total) > 0.01:
            return equal_widths(rows, total)
        return widths

    mode = tables['width_mode']
    if mode == 'equal':
        return safe(equal_widths(rows, total))
    if mode == 'rules':
        widths = rule_widths(rows, tables.get('rules'), total)
        return safe(widths) if widths else safe(equal_widths(rows, total))
    if mode != 'auto':
        raise ConfigError(f"tables.width_mode 只能是 auto / equal / rules，收到 {mode!r}")
    return safe(auto_widths(
        rows,
        total,
        float(tables['minimum_column_cm']),
        float(tables['maximum_weight']),
    ))


# --------------------------------------------------------------------------
# 字体与样式
# --------------------------------------------------------------------------

def set_run_fonts(rpr, config, bold=False, mono=False):
    """给 rPr 设定字体。中文必须写 w:eastAsia，只设 font.name 不管用。"""
    fonts = config['fonts']
    latin = fonts['mono'] if mono else fonts['latin']
    east_asian = fonts['cjk_bold'] if bold else fonts['cjk']

    element = rpr.find(qn('w:rFonts'))
    if element is None:
        element = OxmlElement('w:rFonts')
        rpr.insert(0, element)
    for name in list(element.attrib):
        del element.attrib[name]
    for name in ('ascii', 'hAnsi', 'cs'):
        element.set(qn('w:' + name), latin)
    element.set(qn('w:eastAsia'), east_asian)

    color = rpr.find(qn('w:color'))
    if color is None:
        color = OxmlElement('w:color')
        rpr.append(color)
    color.attrib.clear()
    color.set(qn('w:val'), '000000')

    lang = rpr.find(qn('w:lang'))
    if lang is None:
        lang = OxmlElement('w:lang')
        rpr.append(lang)
    lang.set(qn('w:val'), 'en-US')
    lang.set(qn('w:eastAsia'), 'zh-CN')


def strip_theme(document, config):
    """清掉内置样式的主题色和主题字体，否则 Word 里会冒出蓝色标题。"""
    for style in document.styles:
        if not hasattr(style, 'font'):
            continue
        set_run_fonts(style.element.get_or_add_rPr(), config, bool(style.font.bold))
        style.font.color.rgb = RGBColor(0, 0, 0)


def apply_page(document, config):
    page = config['page']
    section = document.sections[0]
    section.page_width = Cm(page['width_cm'])
    section.page_height = Cm(page['height_cm'])
    for side in ('top', 'bottom', 'left', 'right'):
        setattr(section, f'{side}_margin', Cm(page['margin_cm']))
    section.header_distance = Cm(page['header_distance_cm'])
    section.footer_distance = Cm(page['footer_distance_cm'])
    return section


def _define_style(document, config, name, size, bold=False, line=1.5, before=0, after=0):
    if name not in document.styles:
        document.styles.add_style(name, 1)
    style = document.styles[name]
    style.font.name = config['fonts']['latin']
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.italic = False
    style.font.color.rgb = RGBColor(0, 0, 0)
    set_run_fonts(style.element.get_or_add_rPr(), config, bold)
    fmt = style.paragraph_format
    fmt.line_spacing = Pt(size * line)
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after)
    fmt.widow_control = True
    return style


def define_styles(document, config):
    """建立本报告用到的全部段落样式，并返回常用样式名。"""
    sizes = config['sizes']
    body = config['body']
    headings = config['headings']

    strip_theme(document, config)

    normal = _define_style(document, config, 'Normal', sizes['body'], line=body['line_spacing'])
    normal.paragraph_format.first_line_indent = Pt(body['first_line_indent_pt'])
    normal.paragraph_format.alignment = ALIGNMENTS.get(body['align'], WD_ALIGN_PARAGRAPH.JUSTIFY)

    title = _define_style(document, config, 'Title', sizes['title'], True, 1.3, 0, 10)
    title.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.first_line_indent = Pt(0)

    subtitle = _define_style(document, config, 'Subtitle', sizes['subtitle'], False, 1.3, 0, 16)
    subtitle.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.paragraph_format.first_line_indent = Pt(0)

    for level in (1, 2, 3):
        style = _define_style(
            document, config, f'Heading {level}', sizes[f'h{level}'], True,
            headings['line_spacing'],
            headings['space_before'][level - 1],
            headings['space_after'],
        )
        fmt = style.paragraph_format
        fmt.first_line_indent = Pt(0)
        fmt.keep_with_next = True
        fmt.keep_together = True
        fmt.alignment = WD_ALIGN_PARAGRAPH.LEFT
        # Markdown 里已经手写了「4.1」这类编号，去掉样式自带的自动编号。
        ppr = style.element.find(qn('w:pPr'))
        if ppr is not None:
            numbering = ppr.find(qn('w:numPr'))
            if numbering is not None:
                ppr.remove(numbering)

    caption = _define_style(document, config, 'Caption', sizes['caption'], False, 1.2, 4, 9)
    caption.paragraph_format.first_line_indent = Pt(0)
    caption.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

    table_style = _define_style(document, config, 'Report Table', sizes['table'], False, 1.1)
    table_style.paragraph_format.first_line_indent = Pt(0)

    code_style = _define_style(document, config, 'Report Code', sizes['code'], False, 1.1)
    # 悬挂缩进：首行顶格，折行部分缩进，避免折下来的内容与上一行分不清。
    wrap_indent = config['code']['wrap_indent_cm']
    code_style.paragraph_format.left_indent = Cm(wrap_indent)
    code_style.paragraph_format.first_line_indent = Cm(-wrap_indent)
    code_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    code_style.paragraph_format.widow_control = False
    set_run_fonts(code_style.element.get_or_add_rPr(), config, False,
                  mono=config['code']['font'] == 'mono')

    for name in ('TOC 1', 'TOC 2'):
        style = _define_style(document, config, name, sizes['toc'], False, 1.1, 0, 1)
        style.paragraph_format.first_line_indent = Pt(0)
        style.paragraph_format.left_indent = Cm(0 if name == 'TOC 1' else 0.65)

    return {'normal': 'Normal', 'title': 'Title', 'subtitle': 'Subtitle',
            'caption': 'Caption', 'table': 'Report Table', 'code': 'Report Code'}


def enable_field_updates(document):
    """让 Word 打开文档时提示更新目录和页码。"""
    settings = document.settings.element
    for name in ('hideSpellingErrors', 'hideGrammaticalErrors', 'updateFields'):
        node = OxmlElement('w:' + name)
        node.set(qn('w:val'), 'true')
        settings.append(node)


def add_page_number_footer(section, config):
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.paragraph_format.first_line_indent = Pt(0)
    run = footer.add_run()
    set_run_fonts(run._element.get_or_add_rPr(), config)
    run.font.size = Pt(config['sizes']['footer'])
    begin = OxmlElement('w:fldChar')
    begin.set(qn('w:fldCharType'), 'begin')
    instruction = OxmlElement('w:instrText')
    instruction.set(qn('xml:space'), 'preserve')
    instruction.text = ' PAGE '
    end = OxmlElement('w:fldChar')
    end.set(qn('w:fldCharType'), 'end')
    run._r.extend([begin, instruction, end])


def apply_metadata(document, config):
    metadata = config['metadata']
    if metadata.get('title'):
        document.core_properties.title = metadata['title']
    if metadata.get('subject'):
        document.core_properties.subject = metadata['subject']
    if metadata.get('keywords'):
        document.core_properties.keywords = metadata['keywords']


def blank_paragraph(document, config, space_before=0, space_after=0, line_spacing=1.0):
    """建一个不带首行缩进的空段落，用来做间距。"""
    paragraph = document.add_paragraph()
    fmt = paragraph.paragraph_format
    fmt.first_line_indent = Pt(0)
    fmt.space_before = Pt(space_before)
    fmt.space_after = Pt(space_after)
    fmt.line_spacing = line_spacing
    return paragraph
