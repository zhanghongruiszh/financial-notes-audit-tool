# -*- coding: utf-8 -*-
"""
report_semantics.py — 把某份附注的表格语义解析结果打成一份可读报告。

用法：
    python tools/report_semantics.py <附注.doc> [输出txt] [--examples 6]
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import notes_parser      # noqa: E402
import table_semantics as ts   # noqa: E402


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    path = sys.argv[1]
    out_path = sys.argv[2] if len(sys.argv) > 2 and not sys.argv[2].startswith('--') \
        else '_work/semantics_report.txt'
    n_examples = 6
    if '--examples' in sys.argv:
        n_examples = int(sys.argv[sys.argv.index('--examples') + 1])

    model = notes_parser.parse_notes(path, work_dir='_work/parse')

    lines = [f'源文件: {os.path.basename(path)}']
    for ch in model.chapters:
        ctx = ts.build_chapter_context(model, ch)
        defs = {k: str(v) for k, v in ctx['period_defs'].items()}
        lines.append(f"章<{ch.name}> 期间对照: {defs if defs else '（无）'}")
        if ctx['warnings']:
            lines.extend('    ' + w for w in ctx['warnings'])

    tables = ts.analyze_document(model)
    lines.append('')
    lines.append(f'共解析 {len(tables)} 张表')

    n_head = sum(1 for t in tables if t.header_rows)
    n_total = sum(1 for t in tables if t.total_rows)
    n_role = sum(1 for t in tables if any(c.role for c in t.columns))
    n_period = sum(1 for t in tables if any(c.period for c in t.columns))
    lines.append(f'  有表头        {n_head}')
    lines.append(f'  有合计行      {n_total}')
    lines.append(f'  有角色列      {n_role}')
    lines.append(f'  期间已定位    {n_period}')

    # 角色分布
    from collections import Counter
    role_c = Counter(c.role for t in tables for c in t.columns if c.role)
    lines.append(f'  列角色分布    {dict(role_c)}')

    for tm in tables[:n_examples]:
        lines.append('')
        lines.append(tm.describe())

    lines.append('')
    lines.append('=' * 90)
    lines.append('有合计行且有角色列的表（前 %d 张）：' % n_examples)
    picked = [t for t in tables if t.total_rows and any(c.role for c in t.columns)]
    for tm in picked[:n_examples]:
        lines.append('')
        lines.append(tm.describe())

    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or '.', exist_ok=True)
    open(out_path, 'w', encoding='utf-8').write('\n'.join(lines))
    print(f'written {out_path}  ({len(lines)} lines)')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
