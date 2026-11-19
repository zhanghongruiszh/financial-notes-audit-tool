# -*- coding: utf-8 -*-
"""inspect_check.py — 快速查看生成的成果表结构。"""

import sys

import openpyxl


def main():
    path = sys.argv[1]
    sheet_name = None
    nrows = 30
    for a in sys.argv[2:]:
        if a.isdigit():
            nrows = int(a)
        else:
            sheet_name = a
    wb = openpyxl.load_workbook(path)
    print('工作表:', wb.sheetnames)
    if sheet_name:
        ws = wb[sheet_name]
    else:
        ws = wb['勾稽底稿'] if '勾稽底稿' in wb.sheetnames else wb.active
    print(f'{ws.title} {ws.max_row} 行 × {ws.max_column} 列')
    for r in range(1, min(nrows, ws.max_row) + 1):
        cells = []
        for c in range(1, min(ws.max_column, 8) + 1):
            v = ws.cell(r, c).value
            if v is None:
                continue
            fill = ws.cell(r, c).fill
            mark = ''
            if fill and fill.start_color and fill.start_color.rgb:
                rgb = str(fill.start_color.rgb)
                if rgb.endswith('E2EFDA'):
                    mark = '●绿'
                elif rgb.endswith('FFC7CE'):
                    mark = '●深红'
                elif rgb.endswith('FDE2E2'):
                    mark = '●淡红'
                elif rgb.endswith('FFF2CC'):
                    mark = '●黄'
            cells.append(f'{c}:{v}{mark}')
        if cells:
            print(f'R{r}: ' + ' | '.join(str(x)[:70] for x in cells))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
