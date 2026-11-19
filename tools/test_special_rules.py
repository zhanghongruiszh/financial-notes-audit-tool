# -*- coding: utf-8 -*-
"""
test_special_rules.py — 特殊表加减关系的单元测试。

覆盖两条按模板结构实现的规则：
  R-BLK   固定资产四块式：期末=期初+增加−减少；账面价值=原值−折旧−减值
  R-PROV  坏账准备变动：期末=期初+计提−收回−转销

用轻量假对象构造表格，不依赖 Word。
"""

import sys
import os
from types import SimpleNamespace

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import rules


def make_table(rows, cols, grid):
    """cols: [(表头, 角色, kind)]，第一个通常是项目列（text，不参与计算）。"""
    tm = SimpleNamespace()
    tm.chapter = '六、测试'
    tm.sub_item = '测试科目'
    tm.ref = SimpleNamespace(idx=99, rows=len(grid), cols=len(grid[0]), grid=grid)
    tm.warnings = []
    tm.table_type = 'OTHER'
    tm.table_type_name = ''
    tm.columns = [
        SimpleNamespace(index=i, is_data=(k != 'text'), header_text=h, role=r, kind=k,
                        period=None, period_key=None, path=[h])
        for i, (h, r, k) in enumerate(cols)
    ]
    tm.rows = [SimpleNamespace(row=i + 1, kind=k, label=l) for i, (k, l) in enumerate(rows)]
    return tm


def grid_from(rows):
    out = []
    for r in rows:
        out.append([{'t': ('' if v is None else str(v)), 'm': ' ', 'w': ''} for v in r])
    return out


def test_block_move():
    rows = [
        ('header', '项目'),
        ('section', '① 账面原值'),
        ('detail', '期初余额'),
        ('detail', '本期增加金额'),
        ('detail', '本期减少金额'),
        ('detail', '期末余额'),
        ('section', '② 累计折旧'),
        ('detail', '期初余额'),
        ('detail', '本期增加金额'),
        ('detail', '本期减少金额'),
        ('detail', '期末余额'),
        ('section', '④ 账面价值'),
        ('detail', '期末账面价值'),
        ('detail', '期初账面价值'),
    ]
    data = [
        ['项目', '金额'],
        ['① 账面原值', None],
        ['期初余额', 1000],
        ['本期增加金额', 200],
        ['本期减少金额', 50],
        ['期末余额', 1150],
        ['② 累计折旧', None],
        ['期初余额', 300],
        ['本期增加金额', 100],
        ['本期减少金额', 20],
        ['期末余额', 380],
        ['④ 账面价值', None],
        ['期末账面价值', 770],
        ['期初账面价值', 700],
    ]
    tm = make_table(rows, [('项目', None, 'text'), ('金额', None, 'amount')], grid_from(data))
    checks = []
    rules._rule_block_move(tm, checks)
    by_uid = {tm.ref.idx: tm}
    results = [rules.evaluate(c, tm, by_uid) for c in checks]

    ok = all(r.status == 'pass' for r in results)
    print(f'R-BLK 生成 {len(checks)} 条，结果 {[r.status for r in results]}')
    for c in checks:
        print('   ', c.label)
    return ok


def test_prov_move():
    rows = [
        ('header', '类别'),
        ('detail', '按组合计提'),
        ('total', '合计'),
    ]
    cols = [('类别', None, 'text'), ('期初余额', '余额', 'amount'), ('计提', '金额', 'amount'),
            ('收回或转回', '金额', 'amount'), ('转销或核销', '金额', 'amount'),
            ('期末余额', '余额', 'amount')]
    data = [
        ['类别', '期初余额', '计提', '收回或转回', '转销或核销', '期末余额'],
        ['按组合计提', 1000, 200, 30, 70, 1100],     # 1000+200-30-70 = 1100 ✓
        ['合计', 1000, 200, 30, 70, 1100],
    ]
    tm = make_table(rows, cols, grid_from(data))
    checks = []
    rules._rule_prov_move(tm, checks)
    by_uid = {tm.ref.idx: tm}
    results = [rules.evaluate(c, tm, by_uid) for c in checks]
    ok = all(r.status == 'pass' for r in results)
    print(f'R-PROV 生成 {len(checks)} 条，结果 {[r.status for r in results]}')
    for c in checks:
        print('   ', c.label)

    # 再故意改坏一个数，确认能抓到
    data[1][5] = '1200'
    tm2 = make_table(rows, cols, grid_from(data))
    checks2 = []
    rules._rule_prov_move(tm2, checks2)
    results2 = [rules.evaluate(c, tm2, {tm2.ref.idx: tm2}) for c in checks2]
    caught = any(r.status == 'fail' for r in results2)
    print(f'   改坏后期末差 100 → {[r.status for r in results2]}（应出现 fail）')
    return ok and caught


def test_ecl():
    """三阶段坏账变动表：期末 = 期初 + 期初在本期 + 计提 − 转回 − 转销 − 核销。"""
    rows = [
        ('header', '坏账准备'),
        ('detail', '期初余额'),
        ('detail', '期初余额在本期'),
        ('detail', '--转入第二阶段'),
        ('detail', '--转回第一阶段'),
        ('detail', '本期计提'),
        ('detail', '本期转回'),
        ('detail', '本期转销'),
        ('detail', '本期核销'),
        ('detail', '其他变动'),
        ('detail', '期末余额'),
    ]
    cols = [('坏账准备', None, 'text'), ('第一阶段', None, 'amount'),
            ('第二阶段', None, 'amount'), ('合计', None, 'amount')]
    data = [
        ['坏账准备', '第一阶段', '第二阶段', '合计'],
        ['期初余额', 100, 50, 150],
        ['期初余额在本期', -20, 20, 0],
        ['--转入第二阶段', -20, 20, 0],
        ['--转回第一阶段', 0, 0, 0],
        ['本期计提', 10, 5, 15],
        ['本期转回', 3, 2, 5],
        ['本期转销', 0, 0, 0],
        ['本期核销', 7, 3, 10],
        ['其他变动', 0, 0, 0],
        ['期末余额', 80, 70, 150],   # 100-20+10-3-0-7=80；50+20+5-2-0-3=70
    ]
    tm = make_table(rows, cols, grid_from(data))
    checks = []
    rules._rule_ecl(tm, checks)
    results = [rules.evaluate(c, tm, {tm.ref.idx: tm}) for c in checks]
    ok = bool(checks) and all(r.status == 'pass' for r in results)
    print(f'R-ECL 生成 {len(checks)} 条，结果 {[r.status for r in results]}')
    for c in checks:
        print('   ', c.label)

    # 改坏期末，确认能抓到
    data[-1][1] = '99'
    tm2 = make_table(rows, cols, grid_from(data))
    checks2 = []
    rules._rule_ecl(tm2, checks2)
    results2 = [rules.evaluate(c, tm2, {tm2.ref.idx: tm2}) for c in checks2]
    caught = any(r.status == 'fail' for r in results2)
    print(f'   改坏第一阶段期末 → {[r.status for r in results2]}（应出现 fail）')
    return ok and caught


def main():
    a = test_block_move()
    b = test_prov_move()
    c = test_ecl()
    print()
    print('R-BLK 通过' if a else 'R-BLK 失败')
    print('R-PROV 通过' if b else 'R-PROV 失败')
    print('R-ECL 通过' if c else 'R-ECL 失败')
    return 0 if (a and b and c) else 1


if __name__ == '__main__':
    sys.exit(main())
