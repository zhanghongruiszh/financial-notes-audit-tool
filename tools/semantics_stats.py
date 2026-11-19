# -*- coding: utf-8 -*-
"""
semantics_stats.py — 批量跑 5 份模板，统计表头/合计行/列角色/期间的识别覆盖率。

用法：
    python tools/semantics_stats.py            # 默认跑 附注模板/*.doc
"""

import glob
import os
import sys
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import notes_parser        # noqa: E402
import table_semantics as ts   # noqa: E402


def main():
    files = sorted(glob.glob('附注模板/*.doc')) or sorted(glob.glob('附注模板/*.docx'))
    rows = []
    role_total = Counter()
    warn_samples = []

    for f in files:
        model = notes_parser.parse_notes(f, work_dir='_work/parse')
        tables = ts.analyze_document(model)
        n = len(tables)
        h = sum(1 for t in tables if t.header_rows)
        tr = sum(1 for t in tables if t.total_rows)
        rc = sum(1 for t in tables if any(c.role for c in t.columns))
        pd = sum(1 for t in tables if any(c.period for c in t.columns))
        warns = sum(len(t.warnings) for t in tables)
        for t in tables:
            for c in t.columns:
                if c.role:
                    role_total[c.role] += 1
            if t.warnings and len(warn_samples) < 8:
                warn_samples.append(f'{os.path.basename(f)} | {t.sub_item} 表#{t.ref.idx}: '
                                    f'{t.warnings[0][:60]}')
        rows.append((os.path.basename(f)[:22], n, h, tr, rc, pd, warns))

    print(f'{"模板":<24}{"表数":>6}{"有表头":>7}{"有合计":>7}{"有角色":>7}{"期间定":>7}{"警告":>7}')
    print('-' * 68)
    for r in rows:
        print(f'{r[0]:<24}{r[1]:>6}{r[2]:>7}{r[3]:>7}{r[4]:>7}{r[5]:>7}{r[6]:>7}')

    print('\n列角色合计:', dict(role_total))
    if warn_samples:
        print('\n警告样例:')
        for w in warn_samples:
            print('  -', w)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
