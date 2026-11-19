# -*- coding: utf-8 -*-
"""
mine_rules.py — 用"数值关系"反推表结构规律，输出规则候选。

思路（这是用户给带金额测试数据的本意）：
  空表只有表头，语义要靠猜；有数字的表可以**验证**：
    · 某列 = 其余列之和        → 合计列
    · 某行 = 上方若干行之和    → 合计行
    · 期末 = 期初 + 增 − 减    → 变动表
    · 原值 − 折旧 = 净值       → 分段式资产表
  凡是能验证的，就抽出"表头/行名的关键字组合"，作为规则候选。

用法：
    python tools/mine_rules.py <附注.docx> [--out docs/挖掘结果.md]
"""

import argparse
import os
import sys
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import notes_parser                     # noqa: E402
import table_semantics as ts            # noqa: E402
import rules                            # noqa: E402

TOL = 0.01


def num(grid, r, c):
    return rules.cell_number(grid, r, c) or 0.0


def find_vertical(grid, tm):
    """合计行 = 上方连续明细之和？返回 (合计行, 明细行数)。"""
    hits = []
    for tot in [r for r in tm.rows if r.kind in ('total', 'subtotal')]:
        above = [r for r in tm.rows
                 if r.kind in ('detail', 'group') and r.row < tot.row]
        for col in [c for c in tm.columns if c.is_data]:
            target = num(grid, tot.row, col.index + 1)
            if abs(target) < TOL:
                continue
            acc = sum(num(grid, r.row, col.index + 1) for r in above)
            if abs(acc - target) <= TOL and above:
                hits.append((tot.row, len(above), col.header_text or ''))
                break
    return hits


def find_horizontal(grid, tm):
    """某列 = 其余列之和？返回 [(目标列, 被加的列数)]。"""
    cols = [c for c in tm.columns if c.is_data]
    if len(cols) < 3:
        return []
    hits = []
    rows = [r for r in tm.rows if r.kind in ('detail', 'total', 'subtotal')]
    for i, col in enumerate(cols):
        others = [c for c in cols if c is not col]
        ok = 0
        for r in rows:
            t = num(grid, r.row, col.index + 1)
            acc = sum(num(grid, r.row, c.index + 1) for c in others)
            if abs(t) > TOL and abs(acc - t) <= TOL:
                ok += 1
        if ok >= 2:
            hits.append((col.index + 1, col.header_text or '', ok))
    return hits


def find_asset_block(tm):
    """分段式资产表：原值 / 累计折旧 / 净值 / 减值 / 账面价值。"""
    words = ('原值', '原价', '累计折旧', '累计摊销', '净值', '减值', '账面价值')
    labels = [r.label or '' for r in tm.rows if r.kind in ('total', 'subtotal')]
    kind = [w for w in words if any(w in l for l in labels)]
    return kind if len(kind) >= 3 else None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('notes')
    ap.add_argument('--out', default='docs/挖掘结果.md')
    ap.add_argument('--work', default='_work/mine')
    args = ap.parse_args()

    model = notes_parser.parse_notes(args.notes, work_dir=args.work)
    tables = ts.analyze_document(model)

    vertical = Counter()
    horizontal = Counter()
    asset_blocks = []
    covered = 0

    checks = rules.build_checks(tables)
    covered_tables = {c.table_idx for c in checks}

    for tm in tables:
        v = find_vertical(tm.ref.grid, tm)
        if v:
            vertical[tm.table_type_name or '未分类'] += 1
        h = find_horizontal(tm.ref.grid, tm)
        if h:
            horizontal[tm.table_type_name or '未分类'] += 1
        ab = find_asset_block(tm)
        if ab:
            asset_blocks.append((tm.ref.idx, tm.sub_item, ab))
        if tm.ref.idx in covered_tables:
            covered += 1

    lines = []
    lines.append(f'# 数值挖掘结果（{os.path.basename(args.notes)}）\n')
    lines.append(f'- 表格总数：{len(tables)}')
    lines.append(f'- 已有勾稽规则覆盖：{covered} 张（{covered / max(len(tables), 1):.0%}）')
    lines.append(f'- 数值验证出「合计行=上方之和」：{sum(vertical.values())} 张')
    lines.append(f'- 数值验证出「合计列=其余列之和」：{sum(horizontal.values())} 张')
    lines.append(f'- 分段式资产表：{len(asset_blocks)} 张\n')

    lines.append('## 一、按表类型统计（可验证的纵向合计）\n')
    lines.append('| 表类型 | 表数 |')
    lines.append('|---|---|')
    for k, v in vertical.most_common():
        lines.append(f'| {k} | {v} |')

    lines.append('\n## 二、按表类型统计（可验证的横向合计）\n')
    lines.append('| 表类型 | 表数 |')
    lines.append('|---|---|')
    for k, v in horizontal.most_common():
        lines.append(f'| {k} | {v} |')

    lines.append('\n## 三、分段式资产表（关键字组合）\n')
    for idx, sub, kinds in asset_blocks[:20]:
        lines.append(f'- 表#{idx} {sub}：{" / ".join(kinds)}')

    # 未被覆盖的表 → 规则候选
    uncovered = [t for t in tables if t.ref.idx not in covered_tables]
    lines.append(f'\n## 四、还没有勾稽规则的表（{len(uncovered)} 张）\n')
    for t in uncovered[:30]:
        lines.append(f'- 表#{t.ref.idx} {t.sub_item}　'
                     f'{t.ref.rows}×{t.ref.cols}　{t.table_type_name}')

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    with open(args.out, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines) + '\n')
    print('\n'.join(lines[:12]))
    print(f'\n已写入 {args.out}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
