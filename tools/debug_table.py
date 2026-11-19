# -*- coding: utf-8 -*-
"""
debug_table.py — 打印某张表的网格、语义和生成的校验，用于排查规则误判。

用法：
    python tools/debug_table.py <附注.doc> <表号>
    python tools/debug_table.py <附注.doc> --sub 应付职工薪酬
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import notes_parser        # noqa: E402
import table_semantics as ts   # noqa: E402
import rules               # noqa: E402


def print_grid(tm):
    grid = tm.ref.grid
    print(f'\n网格 {len(grid)} 行：')
    for r, row in enumerate(grid, 1):
        cells = []
        for c, cell in enumerate(row, 1):
            if cell['m'] == ' ':
                cells.append(f'{c}:{cell["t"]}' if cell['t'] else '')
            else:
                cells.append('')
        print(f'  R{r:<3d} ' + ' | '.join(x for x in cells if x))
    print('\n行类型：')
    for r in tm.rows:
        print(f'  R{r.row:<3d} [{r.kind:<8}] {r.label[:50]}')
    print('\n列：')
    for c in tm.columns:
        print(f'  C{c.index + 1:<3d} data={c.is_data} kind={c.kind:<7} '
              f'role={c.role or "-":<5} period={c.period or c.period_key or "-"} '
              f'{" → ".join(c.path)}')


def main():
    path = sys.argv[1]
    model = notes_parser.parse_notes(path, work_dir='_work/parse')
    tables = ts.analyze_document(model)
    by_uid = rules.map_tables(tables)
    checks = rules.build_checks(tables)

    if '--sub' in sys.argv:
        name = sys.argv[sys.argv.index('--sub') + 1]
        picked = [t for t in tables if t.sub_item == name]
    else:
        idx = int(sys.argv[2])
        picked = [t for t in tables if t.ref.idx == idx]

    for tm in picked:
        print('=' * 90)
        print(f'章<{tm.chapter}> {tm.sub_item} 表#{tm.ref.idx} '
              f'{tm.ref.rows}行×{tm.ref.cols}列 表头{tm.header_rows}行')
        print_grid(tm)
        mine = [c for c in checks if c.table_idx == tm.ref.idx
                and c.sub_item == tm.sub_item]
        print(f'\n生成的校验 {len(mine)} 条：')
        for c in mine:
            res = rules.evaluate(c, tm)
            print(f'  [{c.rule_id}] 目标 R{c.target_row}C{c.target_col} = '
                  f'{c.formula_refs()}  → {res.status} '
                  f'(期望 {res.expected}, 实际 {res.actual}, 差 {res.diff})')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
