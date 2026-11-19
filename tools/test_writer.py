# -*- coding: utf-8 -*-
"""
test_writer.py — 跑通「解析 → 规则 → 输出」并生成成果表。

用法：
    python tools/test_writer.py <附注.doc> [输出.xlsx]
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import notes_parser        # noqa: E402
import table_semantics as ts   # noqa: E402
import rules               # noqa: E402
import writer              # noqa: E402


def main():
    path = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else '_work/out/check_test.xlsx'
    os.makedirs(os.path.dirname(os.path.abspath(out)) or '.', exist_ok=True)

    model = notes_parser.parse_notes(path, work_dir='_work/parse')
    tables = ts.analyze_document(model)
    checks = rules.build_checks(tables)
    by_uid = rules.map_tables(tables)

    if '--perturb' in sys.argv:
        # 故意改坏一个数，验证成果表能标红
        first = next((c for c in checks if c.rule_id == 'V-SUM'), None)
        if first:
            tm = by_uid[first.table_uid]
            r, c = first.target_row - 1, first.target_col - 1
            v = rules.cell_number(tm.ref.grid, r + 1, c + 1) or 0.0
            tm.ref.grid[r][c]['t'] = f'{v + 500:,.2f}'
            print(f'已改坏 表#{tm.ref.idx} R{r + 1}C{c + 1}：{v:,.2f} → {v + 500:,.2f}')

    results = rules.evaluate_all(checks, by_uid)

    _, stats = writer.write_workbook(out, model, tables, checks, results)
    print('统计:', stats)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
