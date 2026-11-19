# -*- coding: utf-8 -*-
"""diag_one.py — 单独看某个子项目的跨表匹配细节。"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import notes_parser        # noqa: E402
import table_semantics as ts   # noqa: E402
import report_reader as rr     # noqa: E402
import cross_checks as cc      # noqa: E402
import rules               # noqa: E402


def main():
    notes, report_path, sub_name = sys.argv[1], sys.argv[2], sys.argv[3]
    model = notes_parser.parse_notes(notes, work_dir='_work/parse')
    tables = ts.analyze_document(model)
    report = rr.read_report(report_path)

    group = [t for t in tables if t.sub_item == sub_name]
    print(f'子项目「{sub_name}」下共 {len(group)} 张表')
    tm, tot = cc._pick_main_table(group)
    if tm is None:
        print('  没选出主表')
        return 1
    print(f'  选中 表#{tm.ref.idx}  代表行 R{tot.row} [{tot.kind}] {tot.label}')

    names = cc._split_names(sub_name)
    print(f'  拆分名: {names}')
    sheet, found, matched = cc._match_report(report, '单体', names)
    print(f'  报表命中: {sheet.name if sheet else None} {found} 命中名={matched}')

    hint_keys = []
    if matched and matched != sub_name:
        hint_keys = [matched] + ([matched[-2:]] if len(matched) >= 2 else [])
    print(f'  表头关键字: {hint_keys}')

    for col in tm.columns:
        if not col.is_data:
            continue
        head = (col.header_text or '').replace(' ', '')
        ok = (not hint_keys) or any(k in head for k in hint_keys)
        val = rules.cell_number(tm.ref.grid, tot.row, col.index + 1)
        print(f'    C{col.index + 1} period={col.period} role={col.role} '
              f'kind={col.kind} ok={ok} value={val}  <{col.header_text}>')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
