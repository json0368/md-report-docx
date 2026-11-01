"""检查 docx 里正文与代码块实际使用的字体。

用法：python _check_fonts.py 报告.docx
代码块装箱后位于表格单元格内，必须遍历整个 body 才能看到。
"""
import sys
from collections import Counter
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

for raw in sys.argv[1:]:
    path = Path(raw)
    document = Document(str(path))
    body, code = Counter(), Counter()
    for node in document.element.body.iter(qn('w:p')):
        paragraph = Paragraph(node, document)
        target = code if paragraph.style.name == 'Report Code' else body
        for run in paragraph.runs:
            if not run.text.strip():
                continue
            node = run._r.get_or_add_rPr().find(qn('w:rFonts'))
            if node is None:
                continue
            target[(node.get(qn('w:ascii')), node.get(qn('w:eastAsia')))] += 1
    print(path.name)
    for label, counter in (('正文（含表格）', body), ('代码块', code)):
        print(f'  {label}：')
        if not counter:
            print('      无')
        for (latin, east), count in counter.most_common(6):
            print(f'      ascii={latin:<18} eastAsia={east:<6} {count} 个 run')
