# -*- coding: utf-8 -*-
"""
dump_items.py — 输出报表项目 / 子项清单，供人工核对分类是否合理。

输出 markdown 表格，列：序号 | 报表 | 段 | 行 | 类型 | 项目名称 | 归属 | 金额

用法：
    python tools/dump_items.py <报表.xlsx> [--out docs/报表项目清单.md]
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl                       # noqa: E402
import report_reader as rr            # noqa: E402
import report_items as ri             # noqa: E402

KIND_NAME = {'line': '报表项目', 'sub': '子项', 'total': '合计项',
             'section': '分段标题', 'template': '模板行'}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('report')
    ap.add_argument('--out', default='docs/报表项目清单.md')
    args = ap.parse_args()

    book = rr.read_report(args.report)
    wb = openpyxl.load_workbook(args.report, data_only=True)
    items = ri.ordered_items(book, wb=wb)

    # 各表各段的期间列
    blocks = {}
    for sheet in book.sheets:
        for item_col, cols in (sheet.blocks or {}).items():
            periods = {}
            for col, period in cols.items():
                periods.setdefault(period, col)
            blocks[(sheet.name, item_col)] = sorted(periods.items(), reverse=True)

    def amount(sheet_name, item_col, row):
        cols = blocks.get((sheet_name, item_col)) or []
        if not cols or sheet_name not in wb.sheetnames:
            return None
        ws = wb[sheet_name]
        v = ws.cell(row, cols[0][1]).value
        return float(v) if isinstance(v, (int, float)) else None

    lines = ['# 报表项目 / 子项清单\n']
    lines.append(f'来源：`{os.path.basename(args.report)}`\n')
    counts = {}
    for it in items:
        counts[it.kind] = counts.get(it.kind, 0) + 1
    lines.append('## 汇总\n')
    lines.append('| 类型 | 数量 |')
    lines.append('|---|---|')
    for k, name in KIND_NAME.items():
        lines.append(f'| {name} | {counts.get(k, 0)} |')

    items = [it for it in items if it.kind != 'template']
    cur = None
    for it in items:
        head = f'{it.sheet}（第{it.block}段）'
        if head != cur:
            lines.append(f'\n## {head}\n')
            lines.append('| 行 | 类型 | 项目名称 | 归属 | 金额 |')
            lines.append('|---|---|---|---|---|')
            cur = head
        amt = amount(it.sheet, it.block, it.row)
        amt_s = f'{amt:,.2f}' if amt is not None else ''
        lines.append(f'| {it.row} | {KIND_NAME[it.kind]} | {it.name} | '
                     f'{it.parent or ""} | {amt_s} |')

    os.makedirs(os.path.dirname(args.out) or '.', exist_ok=True)
    with open(args.out, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines) + '\n')

    print(f'共 {len(items)} 行：',
          '、'.join(f'{KIND_NAME[k]} {v}' for k, v in counts.items()))
    print('\n子项清单：')
    for it in items:
        if it.kind == 'sub':
            print(f'  {it.sheet} R{it.row} {it.name} ← 归属：{it.parent}')
    print(f'\n已写入 {args.out}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
