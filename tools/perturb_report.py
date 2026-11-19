# -*- coding: utf-8 -*-
"""
perturb_report.py — 反证：把报表里某个数改坏，确认跨表勾稽能抓到差异。

用法：
    python tools/perturb_report.py <报表.xlsx> <输出.xlsx> [科目名] [期间列字母] [加多少]
"""

import shutil
import sys

import openpyxl

import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def main():
    src, dst = sys.argv[1], sys.argv[2]
    item_name = sys.argv[3] if len(sys.argv) > 3 else '货币资金'
    col = sys.argv[4] if len(sys.argv) > 4 else 'C'
    delta = float(sys.argv[5]) if len(sys.argv) > 5 else 1000.0
    only_sheet = sys.argv[6] if len(sys.argv) > 6 else None

    shutil.copy2(src, dst)
    # data_only=True 才能拿到公式的计算结果（否则读到的是 "=[2]BS!$Y$8" 这种外链公式）
    wb = openpyxl.load_workbook(dst, data_only=True)

    target = None
    for name in wb.sheetnames:
        if only_sheet and name != only_sheet:
            continue
        ws = wb[name]
        for r in range(1, min(ws.max_row, 120) + 1):
            v = ws.cell(r, 1).value
            if isinstance(v, str) and v.strip().replace(' ', '') == item_name:
                target = (name, r)
                break
        if target:
            break

    if not target:
        print(f'没找到科目「{item_name}」')
        return 1

    sheet, row = target
    cell = wb[sheet][f'{col}{row}']
    old = cell.value
    try:
        base = float(old or 0)
    except (TypeError, ValueError):
        print(f'该格不是数字（{old!r}），无法改')
        return 1
    cell.value = base + delta
    wb.save(dst)
    print(f'已改坏：{sheet}!{col}{row}（{item_name}）{old} → {cell.value}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
