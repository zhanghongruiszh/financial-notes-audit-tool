# -*- coding: utf-8 -*-
"""diag_cross.py — 诊断每个子项目的跨表勾稽为什么没生成。"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import notes_parser        # noqa: E402
import table_semantics as ts   # noqa: E402
import report_reader as rr     # noqa: E402
import cross_checks        # noqa: E402


def main():
    notes, report_path = sys.argv[1], sys.argv[2]
    model = notes_parser.parse_notes(notes, work_dir='_work/parse')
    tables = ts.analyze_document(model)
    report = rr.read_report(report_path)

    by_sub = {}
    for tm in tables:
        by_sub.setdefault((tm.chapter, tm.sub_item), []).append(tm)

    print(f"{'子项目':<20} {'主表':>6} {'期间数':>6} {'报表命中':<28} 说明")
    print('-' * 100)
    for ch in model.chapters:
        for sub in ch.sub_items:
            group = by_sub.get((ch.name, sub.name), [])
            tm, tot = cross_checks._pick_main_table(group)
            if tm is None:
                print(f'{sub.name:<20} {"-":>6} {"-":>6} {"-":<28} 没有带合计行且期间可定位的表')
                continue
            periods = {c.period for c in tm.columns if c.is_data and c.period}
            hits = []
            for kind in ('资产负债表', '利润表', '现金流量表'):
                sheet = report.find(ch.scope, kind) or report.find('单体', kind)
                if sheet and rr.find_item(sheet, sub.name):
                    hits.append(sheet.name)
            print(f'{sub.name:<20} #{tm.ref.idx:<5} {len(periods):>6} '
                  f'{(hits[0] if hits else "未命中"):<28} '
                  f'{"表内有合计行" if tot else ""}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
