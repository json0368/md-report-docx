"""检查 docx 里代码块每个 run 的实际颜色。

用法：python _check_colors.py 报告.docx [关键字]
不给关键字就列出所有代码块的颜色分布；给了就只打印包含它的代码块逐行明细。
"""
import sys
from collections import Counter
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph


def code_paragraphs(document):
    for node in document.element.body.iter(qn('w:p')):
        paragraph = Paragraph(node, document)
        if paragraph.style.name == 'Report Code':
            yield paragraph


def run_color(run):
    node = run._r.get_or_add_rPr().find(qn('w:color'))
    return node.get(qn('w:val')) if node is not None else None


for raw in sys.argv[1:2]:
    document = Document(str(Path(raw)))
    keyword = sys.argv[2] if len(sys.argv) > 2 else None
    print(Path(raw).name)
    if keyword is None:
        counter = Counter()
        for paragraph in code_paragraphs(document):
            for run in paragraph.runs:
                if run.text.strip():
                    counter[run_color(run)] += 1
        for color, count in counter.most_common():
            print(f'    {color}  {count} 个 run')
    else:
        for paragraph in code_paragraphs(document):
            if keyword not in paragraph.text:
                continue
            print(f'    段落：{paragraph.text}')
            for run in paragraph.runs:
                if run.text.strip():
                    print(f'        {run_color(run)}  {run.text!r}')
            print()
