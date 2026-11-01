"""探针：逐个语言标记实际跑一遍分词，看能不能真正着色。

判定标准：除了 text 之外还出现了别的配色类别，才算这个语言有语法着色。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / 'scripts'))

from pygments import lex
from pygments.lexers import get_all_lexers, get_lexer_by_name
from pygments.util import ClassNotFound

import code_highlight

SAMPLES = [
    ('sql', 'SELECT id FROM t WHERE a = 1; -- 注释'),
    ('mysql', 'SELECT `id` FROM `t` LIMIT 10;'),
    ('java', 'public class A { private int x = 1; // 注释\n}'),
    ('xml', '<?xml version="1.0"?><insert id="a"><if test="x">1</if></insert>'),
    ('html', '<div class="a" id="b">text</div>'),
    ('javascript', 'async function run(a) { const x = 1; // 注释\nreturn a; }'),
    ('js', 'const x = 1; let y = "s";'),
    ('typescript', 'const x: number = 1;'),
    ('jsx', 'const A = () => <div className="a"/>;'),
    ('tsx', 'const A = (): JSX.Element => <div/>;'),
    ('vue', '<template><div>{{ x }}</div></template>'),
    ('python', 'def f(x):\n    return x + 1  # 注释'),
    ('c', 'int main(void) { return 0; }'),
    ('cpp', '#include <vector>\nint main() { return 0; }'),
    ('csharp', 'public int X { get; set; }'),
    ('go', 'func main() { fmt.Println("hi") }'),
    ('rust', 'fn main() { let x = 1; }'),
    ('php', '<?php echo "hi"; ?>'),
    ('ruby', 'def f; 1; end'),
    ('kotlin', 'fun main() { println("hi") }'),
    ('scala', 'object A { def f = 1 }'),
    ('swift', 'let x = 1'),
    ('dart', 'void main() {}'),
    ('objectivec', 'int main() { return 0; }'),
    ('bash', 'if [ -f a ]; then echo "yes"; fi'),
    ('sh', 'mvn -DskipTests package && java -jar a.jar'),
    ('shell', 'ls -la | grep a'),
    ('console', '$ ls -la'),
    ('powershell', 'Get-ChildItem -Recurse -Force'),
    ('batch', 'echo off\nif exist a del a'),
    ('json', '{"a": 1, "b": true}'),
    ('yaml', 'key: value\nlist:\n  - a'),
    ('toml', '[a]\nb = 1'),
    ('ini', '[section]\nkey=value'),
    ('properties', 'a.b=c'),
    ('docker', 'FROM ubuntu:24.04\nRUN apt-get update'),
    ('makefile', 'all:\n\tgcc -o a a.c'),
    ('cmake', 'project(A)\nadd_executable(a a.c)'),
    ('nginx', 'server { listen 80; }'),
    ('apache', 'RewriteEngine On'),
    ('http', 'GET / HTTP/1.1\nHost: a.com'),
    ('css', 'body { color: red; }'),
    ('scss', '$a: red;\nbody { color: $a; }'),
    ('less', '@a: red;\nbody { color: @a; }'),
    ('markdown', '# Title\n**bold**'),
    ('rst', 'Title\n====='),
    ('latex', '\\documentclass{article}'),
    ('tex', '\\begin{document}\\end{document}'),
    ('diff', '+added\n-removed'),
    ('csv', 'a,b,c\n1,2,3'),
    ('text', 'plain text here'),
    ('plaintext', 'plain text here'),
    ('log', '2026-01-01 INFO started'),
    ('protobuf', 'message A { int32 x = 1; }'),
    ('graphql', 'query { user { id } }'),
    ('regex', '^a.*b$'),
    ('lua', 'local x = 1'),
    ('perl', 'my $x = 1;'),
    ('r', 'x <- c(1,2,3)'),
    ('matlab', 'x = 1:10;'),
    ('julia', 'x = 1'),
    ('fortran', 'program a\nend'),
    ('pascal', 'begin end.'),
    ('vbnet', 'Dim x As Integer'),
    ('haskell', 'main = putStrLn "hi"'),
    ('erlang', '-module(a).'),
    ('clojure', '(def x 1)'),
    ('groovy', 'def x = 1'),
    ('asm', 'mov eax, 1'),
    ('nasm', 'mov eax, 1'),
    ('verilog', 'module a; endmodule'),
    ('vhdl', 'entity a is end;'),
    ('solidity', 'contract A { uint x; }'),
    ('abap', 'WRITE hello.'),
]

print(f'pygments 自带词法器：{len(list(get_all_lexers()))} 个\n')

colored, plain, unknown = [], [], []
for tag, sample in SAMPLES:
    name = code_highlight.LEXER_ALIASES.get(tag, tag)
    try:
        lexer = get_lexer_by_name(name, stripnl=False, ensurenl=False)
    except ClassNotFound:
        unknown.append(tag)
        continue
    kinds = set()
    lines, fell_back = code_highlight.lex_lines(sample, tag, None)
    for tokens in lines:
        for _, kind, _ in tokens:
            kinds.add(kind)
    kinds.discard('text')
    (colored if kinds else plain).append((tag, sorted(kinds)))

print(f'能真正着色：{len(colored)} 个')
for tag, kinds in colored:
    print(f'    {tag:<14} {", ".join(kinds)}')

print(f'\n只当纯文本（词法器存在但没分色）：{len(plain)} 个')
print('    ' + ', '.join(tag for tag, _ in plain))

print(f'\n词法器不存在（退回纯文本，不报错）：{len(unknown)} 个')
print('    ' + (', '.join(unknown) if unknown else '无'))
