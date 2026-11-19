# -*- coding: utf-8 -*-
"""test_yellow.py — 验证"结构认不出的表标黄"这条路径确实生效。"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import notes_parser        # noqa: E402
import table_semantics as ts   # noqa: E402
import rules               # noqa: E402
import writer              # noqa: E402


def main():
    path = sys.argv[1]
    out = sys.argv[2]
    model = notes_parser.parse_notes(path, work_dir='_work/parse')
    tables = ts.analyze_document(model)
    checks = rules.build_checks(tables)
    results = rules.evaluate_all(checks, rules.map_tables(tables))

    # 人为给第一张表加一条"结构没认出来"的警告
    tables[0].warnings.insert(0, '没识别出表头（前几行都不像列名）')

    _, stats = writer.write_workbook(out, model, tables, checks, results)
    print('结构存疑表数 =', stats['unclear_tables'])

    import openpyxl
    wb = openpyxl.load_workbook(out)
    ws = wb['勾稽底稿']
    for r in range(1, 12):
        cell = ws.cell(r, 1)
        rgb = str(cell.fill.start_color.rgb or '') if cell.fill else ''
        if rgb.endswith('FFF2CC'):
            print(f'第 {r} 行标黄：{cell.value}')
            return 0
    print('没找到黄色单元格')
    return 1


if __name__ == '__main__':
    raise SystemExit(main())
