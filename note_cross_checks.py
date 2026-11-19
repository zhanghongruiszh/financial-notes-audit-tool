# -*- coding: utf-8 -*-
"""
note_cross_checks.py — 附注内部跨表、跨子项目的一致性核对。

两类
----
**硬勾稽（跨子项目，口径明确）**
    现金流量表补充资料「现金的期末余额」= 货币资金「合计」同期
    现金流量表补充资料「现金的期初余额」= 货币资金「合计」上一期

**参考（同一子项目内跨表）**
    同一子项目、同一类型、同列数的两张表，同期间同角色的合计数应当一致。
    限定"同类型 + 同列数"是为了避免把口径不同的表（如账龄表 vs 组合计提表，
    后者只覆盖组合部分）拿来硬比。仍标为"参考"，供人工判断。
"""

import datetime
from dataclasses import dataclass
from typing import List, Optional

import rules


@dataclass
class NoteXCheck:
    kind: str                      # 硬勾稽 / 参考
    period: Optional[datetime.date]
    role: str = ''
    sub_a: str = ''
    table_a: int = 0
    row_a: int = 0
    value_a: Optional[float] = None
    sub_b: str = ''
    table_b: int = 0
    row_b: int = 0
    value_b: Optional[float] = None
    diff: Optional[float] = None
    status: str = 'unmatched'      # pass / fail / unmatched
    note: str = ''


def _find_row(tm, contains):
    for r in tm.rows:
        if r.label and contains in r.label:
            return r
    return None


def _total_row(tm):
    for r in tm.rows:
        if r.kind == 'total':
            return r
    return None


def _col_of_period(tm, period, roles=('余额', '金额', '价值')):
    for c in tm.columns:
        if not c.is_data or c.period != period:
            continue
        if c.role in roles or c.role is None:
            return c
    return None


def _periods_desc(tm):
    return sorted({c.period for c in tm.columns if c.is_data and c.period},
                  reverse=True)


def _hard_links(model, tables):
    """跨子项目的固定口径勾稽。"""
    out = []
    cf_tables = [t for t in tables if '现金流量表' in (t.sub_item or '')]
    cash_tables = [t for t in tables if (t.sub_item or '').strip() == '货币资金']
    if not cf_tables or not cash_tables:
        return out

    cash_tm = cash_tables[0]
    cash_total = _total_row(cash_tm)
    if cash_total is None:
        return out
    cash_periods = _periods_desc(cash_tm)

    for cf_tm in cf_tables:
        for label, offset, note in (('现金的期末余额', 0, '期末余额 vs 货币资金合计'),
                                    ('现金的期初余额', 1, '期初余额 vs 货币资金上期合计')):
            row = _find_row(cf_tm, label)
            if row is None:
                continue
            for col in cf_tm.columns:
                if not col.is_data or not col.period:
                    continue
                val_a = rules.cell_number(cf_tm.ref.grid, row.row, col.index + 1)
                if val_a is None:
                    continue
                idx = cash_periods.index(col.period) if col.period in cash_periods else -1
                if idx < 0 or idx + offset >= len(cash_periods):
                    continue
                target_period = cash_periods[idx + offset]
                ccol = _col_of_period(cash_tm, target_period)
                if ccol is None:
                    continue
                val_b = rules.cell_number(cash_tm.ref.grid, cash_total.row, ccol.index + 1)
                if val_b is None:
                    continue

                chk = NoteXCheck(
                    kind='硬勾稽', period=col.period, role='金额',
                    sub_a=cf_tm.sub_item, table_a=cf_tm.ref.idx, row_a=row.row,
                    value_a=val_a,
                    sub_b='货币资金', table_b=cash_tm.ref.idx, row_b=cash_total.row,
                    value_b=val_b,
                    note=f'{note}（{label} @{col.period}）',
                )
                chk.diff = val_a - val_b
                chk.status = 'pass' if abs(chk.diff) <= rules.TOLERANCE else 'fail'
                out.append(chk)
    return out


def _detail_labels(tm):
    return tuple(sorted((r.label or '').strip() for r in tm.rows
                        if r.kind == 'detail' and r.label))


def _reference(model, tables):
    """同一子项目内、**行标签完全一致**的两张表，合计行同期同角色比一比。

    只用"同类型同列数"不够——同一子项目里经常并列放着分类表、账龄表、
    变动明细表，它们口径本来就不同（比如"其他应付款"下面既有按性质分类，
    也有按账龄分类）。所以再加一条硬条件：**明细行标签完全一样**，
    这样才是"同一张表的两份呈现"。
    """
    groups = {}
    for tm in tables:
        tot = _total_row(tm)
        if tot is None:
            continue
        if getattr(tm, 'table_type', '') in ('', 'OTHER'):
            continue
        labels = _detail_labels(tm)
        if not labels:
            continue
        key = (tm.sub_item, tm.table_type, tm.ref.cols, labels)
        groups.setdefault(key, []).append((tm, tot))

    out = []
    for (sub, ttype, cols, labels), items in groups.items():
        if len(items) < 2:
            continue
        base_tm, base_tot = items[0]
        for tm, tot in items[1:]:
            for period in sorted({c.period for c in base_tm.columns if c.is_data and c.period}):
                ca = _col_of_period(base_tm, period)
                cb = _col_of_period(tm, period)
                if ca is None or cb is None:
                    continue
                va = rules.cell_number(base_tm.ref.grid, base_tot.row, ca.index + 1)
                vb = rules.cell_number(tm.ref.grid, tot.row, cb.index + 1)
                if va is None or vb is None:
                    continue
                if abs(va) < 0.005 and abs(vb) < 0.005:
                    continue
                chk = NoteXCheck(
                    kind='参考', period=period,
                    role=ca.role or '',
                    sub_a=sub, table_a=base_tm.ref.idx, row_a=base_tot.row, value_a=va,
                    sub_b=sub, table_b=tm.ref.idx, row_b=tot.row, value_b=vb,
                    note=f'同子项目下两表（类型 {ttype}）合计不一致，可能口径不同，供参考',
                )
                chk.diff = va - vb
                chk.status = 'pass' if abs(chk.diff) <= rules.TOLERANCE else 'fail'
                out.append(chk)
    return out


def build(model, tables):
    out = _hard_links(model, tables) + _reference(model, tables)
    out.sort(key=lambda x: (x.kind != '硬勾稽', x.sub_a, x.table_a, x.table_b))
    return out


def summarize(checks):
    from collections import Counter
    return dict(Counter(c.status for c in checks))
