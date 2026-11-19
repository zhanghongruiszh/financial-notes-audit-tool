# -*- coding: utf-8 -*-
"""debug_report.py — 查看报表解析结果。"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl      # noqa: E402
import report_reader as rr   # noqa: E402


def main():
    path = sys.argv[1]
    book = rr.read_report(path)
    for s in book.sheets:
        print('=' * 80)
        print(f'[{s.name}] scope={s.scope} kind={s.kind} header_row={s.header_row}')
        print('  期间列:', {c: str(d) for c, d in s.period_columns()})
        print(f'  科目 {len(s.items)} 个，前 12 个：')
        for k, v in list(s.items.items())[:12]:
            print(f'    R{v[0]:<4d} C{v[1]:<3d} {k}')

    if len(sys.argv) > 2:
        target = sys.argv[2]
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb[target] if target in wb.sheetnames else wb.active
        print('=' * 80)
        print(f'[{target}] 前 20 行原始内容：')
        for r in range(1, min(20, ws.max_row) + 1):
            vals = []
            for c in range(1, min(ws.max_column, 8) + 1):
                v = ws.cell(r, c).value
                if v is not None:
                    vals.append(f'{c}:{v}')
            print(f'  R{r:<3d} ' + ' | '.join(str(x)[:40] for x in vals))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
