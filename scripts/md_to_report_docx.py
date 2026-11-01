"""把 Markdown 报告转换成中文课程报告排版的 Word 文档。

用法：
    python md_to_report_docx.py 报告.md
    python md_to_report_docx.py 报告.md --out 输出.docx --check
    python md_to_report_docx.py --print-defaults > 报告.json

在 Markdown 旁边放一个同名的 .toml / .json / .yaml 即可覆盖任意排版参数，
也可以显式用 --config 指定。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import audit_report
import code_highlight
import report_builder
import report_style as rs


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description='把 Markdown 报告转换成中文课程报告排版的 Word 文档。',
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument('markdown', nargs='?', help='输入的 Markdown 文件')
    parser.add_argument('-o', '--out', help='输出的 .docx 路径，默认与输入同名')
    parser.add_argument('-c', '--config', help='排版配置文件（.toml / .json / .yaml）')
    parser.add_argument('--check', action='store_true',
                        help='转换后自检：表宽、图片宽度、颜色、代码块、正文文本')
    parser.add_argument('--pdf', help='配合 --check，额外检查 PDF 里有没有文字溢出页面')
    parser.add_argument('--print-defaults', action='store_true',
                        help='打印全部默认配置（JSON）后退出')
    parser.add_argument('--quiet', action='store_true', help='只输出错误')
    return parser.parse_args(argv)


def report(message, quiet):
    if not quiet:
        print(message, flush=True)


def main(argv=None):
    args = parse_args(argv)

    if args.print_defaults:
        print(json.dumps(rs.DEFAULTS, ensure_ascii=False, indent=2))
        return 0

    if not args.markdown:
        print('错误：请给出要转换的 Markdown 文件。', file=sys.stderr)
        return 2

    source = Path(args.markdown).resolve()
    if not source.is_file():
        print(f'错误：找不到 {source}', file=sys.stderr)
        return 2

    target = Path(args.out).resolve() if args.out else source.with_suffix('.docx')

    try:
        config, config_path = rs.load_config(source, args.config)
    except rs.ConfigError as error:
        print(f'错误：{error}', file=sys.stderr)
        return 2

    report(f'输入  {source}', args.quiet)
    report(f'配置  {config_path if config_path else "内置默认值"}', args.quiet)
    report(f'输出  {target}', args.quiet)

    document, builder = report_builder.build_document(source, config)
    report(f'转换  表格 {len(document.tables)} 张，图片 {len(document.inline_shapes)} 张', args.quiet)

    before_text = audit_report.body_text(document)
    result = code_highlight.frame_code_blocks(document, source, config)
    report(f'代码  {result["framed"]} 个代码块已装箱上色', args.quiet)
    if result.get('plain_fallback'):
        report(f'      {result["plain_fallback"]} 个未着色（词法器会改动内容，已退回纯文本）', args.quiet)

    for warning in builder.warnings:
        print(f'警告：{warning}', file=sys.stderr)
    if result.get('mismatched'):
        print(f'警告：第 {result["mismatched"]} 个代码块内容与 Markdown 有出入。', file=sys.stderr)

    target.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(target))
    report(f'完成  {target.stat().st_size / 1024:.0f} KB', args.quiet)

    if args.check:
        audit = audit_report.audit(document, before_text, source, config, args.pdf)
        summary = {
            '表格': audit['structure']['dataTables'],
            '代码块': audit['structure']['codeTables'],
            '图片': audit['structure']['images'],
            '正文段落': audit['text']['paragraphs'],
            '配色': audit['colors']['colors'],
        }
        report('自检  ' + json.dumps(summary, ensure_ascii=False), args.quiet)
        for warning in audit['warnings']:
            print(f'警告：{warning}', file=sys.stderr)
        if audit['errors']:
            for error in audit['errors']:
                print(f'错误：{error}', file=sys.stderr)
            return 1
        report('自检通过，未发现问题。', args.quiet)

    return 0


if __name__ == '__main__':
    raise SystemExit(main())
