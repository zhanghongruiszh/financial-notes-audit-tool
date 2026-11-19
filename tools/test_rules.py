# -*- coding: utf-8 -*-
"""
test_rules.py — 跑一遍规则库，看生成多少条勾稽、多少平、多少有差。

用法：
    python tools/test_rules.py <附注.doc>
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import notes_parser        # noqa: E402
import table_semantics as ts   # noqa: E402
import rules               # noqa: E402


def perturb_check(path):
    """反证：故意把一个数改坏，确认规则能抓出来。"""
    import copy
    model = notes_parser.parse_notes(path, work_dir='_work/parse')
    tables = ts.analyze_document(model)
    checks = rules.build_checks(tables)
    by_uid = rules.map_tables(tables)
    before = rules.evaluate_all(checks, by_uid)
    fails_before = sum(1 for r in before if r.status == 'fail')

    # 找第一条 V-SUM 校验，把它目标格改大 100
    target = next((r for r in before if r.item.rule_id == 'V-SUM'), None)
    if target is None:
        print('没有可用的 V-SUM 校验，无法反证')
        return 1
    tm = by_uid[target.item.table_uid]
    r, c = target.item.target_row - 1, target.item.target_col - 1
    cell = tm.ref.grid[r][c]
    old = cell['t']
    try:
        val = rules.cell_number(tm.ref.grid, r + 1, c + 1)
        cell['t'] = f'{val + 100:,.2f}'
    except Exception:
        print('改数失败')
        return 1

    after = rules.evaluate_all(checks, by_uid)
    fails_after = sum(1 for r2 in after if r2.status == 'fail')
    hit = next((r2 for r2 in after if r2.status == 'fail'), None)

    print(f'反证：把 表#{tm.ref.idx} {tm.sub_item} R{r + 1}C{c + 1} 从 {old} 改成 {cell["t"]}')
    print(f'  改前有差异 {fails_before} 条 → 改后 {fails_after} 条')
    if hit:
        print(f'  抓到：{hit.item.label[:50]}  期望 {hit.expected:,.2f} '
              f'实际 {hit.actual:,.2f} 差额 {hit.diff:,.2f}')
    return 0 if (fails_after > fails_before and hit) else 1


def main():
    path = sys.argv[1]
    if '--perturb' in sys.argv:
        return perturb_check(path)

    model = notes_parser.parse_notes(path, work_dir='_work/parse')
    tables = ts.analyze_document(model)
    checks = rules.build_checks(tables)
    by_uid = rules.map_tables(tables)
    results = rules.evaluate_all(checks, by_uid)

    print(f'表 {len(tables)} 张   勾稽校验 {len(checks)} 条')
    print('结果分布:', rules.summarize(results))

    from collections import Counter
    rule_c = Counter(r.item.rule_id for r in results)
    print('按规则:', dict(rule_c))

    fails = [r for r in results if r.status == 'fail']
    print(f'\n有差异 {len(fails)} 条，前 12 条：')
    for r in fails[:12]:
        print(f'  [{r.item.rule_id}] {r.item.chapter} / {r.item.sub_item} '
              f'表#{r.item.table_idx} {r.item.label[:44]}')
        print(f'        期望 {r.expected:,.2f}  实际 {r.actual:,.2f}  差额 {r.diff:,.2f}')

    skips = [r for r in results if r.status == 'skip']
    if skips:
        print(f'\n跳过 {len(skips)} 条，原因示例：')
        for r in skips[:5]:
            print(f'  [{r.item.rule_id}] {r.item.sub_item} 表#{r.item.table_idx}: {r.note}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
