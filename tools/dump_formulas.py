# -*- coding: utf-8 -*-
"""
dump_formulas.py — 勾稽公式清单 & 表头覆盖率。

回答两个问题：
  1. 每个科目具体用的是什么勾稽公式？
  2. 表的每一列（表头）是否都被勾稽覆盖了？哪些列没人管？

输出：
  docs/勾稽公式清单.md   逐表逐列列出公式
  终端打印覆盖率汇总

用法：
    python tools/dump_formulas.py <附注.docx> [--out docs/勾稽公式清单.md]
"""

import argparse
import os
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import notes_parser                     # noqa: E402
import table_semantics as ts            # noqa: E402
import rules                            # noqa: E402

RULE_DESC = {
    'V-SUM': '合计行 = 上方明细之和',
    'V-GRP': '「其中」父行 = 子项之和',
    'H-SUM': '合计列 = 其余列之和',
    'R-NET': '账面价值 = 账面余额 − 准备',
    'R-NBV': '账面价值 = 原值 − 折旧 − 减值',
    'R-ABLK': '净值 = 原值 − 折旧；价值 = 净值 − 减值',
    'R-MOVE': '期末 = 期初 + 增加 − 减少',
    'R-TAX': '税后 = 税前 − 所得税',
    'R-PCT': '占比 = 行金额 ÷ 合计',
    'R-SUMD': '合计在上：明细之和 = 合计',
    'R-BLK': '四块式变动表',
    'R-ECL': '三阶段变动：期末 = 期初 + 各项变动',
    'R-PROV': '坏账准备变动',
}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('notes')
    ap.add_argument('--out', default='docs/勾稽公式清单.md')
    ap.add_argument('--work', default='_work/fml')
    args = ap.parse_args()

    model = notes_parser.parse_notes(args.notes, work_dir=args.work)
    tables = ts.analyze_document(model)
    checks = rules.build_checks(tables)

    by_table = defaultdict(list)
    for c in checks:
        by_table[c.table_idx].append(c)

    total_cols = covered_cols = 0
    lines = ['# 勾稽公式清单\n']
    lines.append(f'来源：`{os.path.basename(args.notes)}`\n')

    body = []
    for tm in tables:
        cols = [c for c in tm.columns if c.is_data]
        if not cols:
            continue
        cs = by_table.get(tm.ref.idx, [])
        used = {c.target_col for c in cs}
        total_cols += len(cols)
        covered_cols += len([c for c in cols if c.index + 1 in used])

        body.append(f'\n## 表#{tm.ref.idx} {tm.sub_item} ')
        body.append(f'{tm.ref.rows}行×{tm.ref.cols}列　{tm.table_type_name or ""}\n')
        body.append('| 列 | 表头 | 勾稽公式 |')
        body.append('|---|---|---|')
        for col in cols:
            idx = col.index + 1
            hits = [c for c in cs if c.target_col == idx]
            if hits:
                desc = '；'.join(
                    f'{h.rule_id}：{h.label}' for h in hits)
            else:
                desc = '**（无勾稽）**'
            body.append(f'| {idx} | {col.header_text or ""} | {desc} |')

    pct = covered_cols / max(total_cols, 1)
    lines.append(f'- 表格数：{len(tables)}')
    lines.append(f'- 数据列总数：{total_cols}')
    lines.append(f'- 已被勾稽覆盖的列：{covered_cols}（{pct:.0%}）\n')

    unused = [t for t in tables
              if not any(c.table_idx == t.ref.idx for c in checks)]
    lines.append(f'## 完全没有勾稽的表（{len(unused)} 张）\n')
    for t in unused:
        lines.append(f'- 表#{t.ref.idx} {t.sub_item}　'
                     f'{t.ref.rows}×{t.ref.cols}　{t.table_type_name}')
    lines.append('\n## 逐表逐列公式\n')
    lines.extend(body)

    os.makedirs(os.path.dirname(args.out) or '.', exist_ok=True)
    with open(args.out, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines) + '\n')

    print(f'表格 {len(tables)} 张，数据列 {total_cols}，'
          f'有勾稽的列 {covered_cols}（{pct:.0%}）')
    print(f'完全没有勾稽的表：{len(unused)} 张')
    for t in unused[:8]:
        print(f'  #{t.ref.idx} {t.sub_item} {t.ref.rows}×{t.ref.cols}')
    print(f'已写入 {args.out}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
