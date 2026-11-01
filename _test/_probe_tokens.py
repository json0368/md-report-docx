"""探针：打印某语言的词法输出，并核对拼接结果是否与原文一致。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / 'scripts'))

from pygments import lex
from pygments.lexers import get_lexer_by_name
from pygments.util import ClassNotFound

import code_highlight

SAMPLES = [
    ('diff', '- SELECT * FROM orders WHERE user_id = 1;\n+ SELECT id, created_at FROM orders WHERE user_id = 1;'),
    ('console', '$ mvn -DskipTests package\nBUILD SUCCESS'),
    ('bash', '$ mvn -DskipTests package'),
]

for language, sample in SAMPLES:
    print(f'=== {language} ===')
    name = code_highlight.LEXER_ALIASES.get(language, language)
    try:
        lexer = get_lexer_by_name(name, stripnl=False, ensurenl=False)
    except ClassNotFound:
        print('    词法器不存在')
        continue
    pieces = []
    for token_type, value in lex(sample, lexer):
        if not value:
            continue
        kind, bold = code_highlight.token_kind(token_type, value, language)
        pieces.append(value)
        print(f'    {str(token_type):<34} -> {kind:<10} {value!r}')
    rebuilt = ''.join(pieces)
    print(f'    原文与重建一致：{rebuilt == sample}')
    if rebuilt != sample:
        print(f'      原文 {sample!r}')
        print(f'      重建 {rebuilt!r}')
    print()
