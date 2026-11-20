# -*- coding: utf-8 -*-
"""
matcher.py — 报表项目/子项 ↔ Word 附注 的通用匹配层。

严格按《需求总表》第七节实现，不含任何项目专属的名字：

  报表项目 → 匹配 Word 小标题（子项目标题）
      有金额 + 能匹配   → 之后与该标题下的**合计数**做数字勾稽
      0/空  + 不能匹配  → 正常没有这项
      0/空  + 能匹配    → 按 0 做勾稽
      有金额 + 不能匹配 → 漏项（进未披露，人工判断）

  子项 → 匹配所属报表项目标题下的**表格行**
      归属项目：资产负债表/利润表往上找、现金流量表往下找（见 report_items）
      再在该项目的表里按行名找

匹配手段（按优先级，全部通用）：
  1. 名称：去括号、去空白后互相包含
  2. 报表附注号 ↔ Word 小节编号（前提：人工核对过编号）
  3. 别名表 / 语义映射（规则库，可编辑）
"""

import re

import report_items as ri


def _norm(text):
    s = re.sub(r'[（(][^）)]*[）)]', '', str(text or ''))
    s = re.sub(r'\s+', '', s)
    for pre in ('其中', '加：', '加:', '减：', '减:'):
        if s.startswith(pre):
            s = s[len(pre):]
    return s


class Match:
    __slots__ = ('item', 'sub_item', 'table_idx', 'row', 'cols', 'how')

    def __init__(self, item, sub_item, table_idx, row, cols, how):
        self.item = item          # ReportItem
        self.sub_item = sub_item  # 命中的 Word 小节名
        self.table_idx = table_idx
        self.row = row            # 表内行号（报表项目取合计数行）
        self.cols = cols          # {期间: 列号}
        self.how = how            # '名称' / '编号' / '别名'

    def __repr__(self):
        return (f'<{self.item.kind} {self.item.name} → {self.sub_item}'
                f' T{self.table_idx} R{self.row} [{self.how}]>')


def _tables_by_sub(model):
    out = {}
    for ch in model.chapters:
        for sub in ch.sub_items:
            for tm in sub.tables:
                out.setdefault(sub.name, []).append(tm)
    return out


def _pick_total_row(tm):
    """取该表的"合计数"行：优先带「合计」，其次带「账面价值/净值」，再次末行。"""
    rows = [r for r in tm.rows if r.kind in ('total', 'subtotal')]
    for r in rows:
        if '合计' in (r.label or ''):
            return r
    for r in rows:
        if any(k in (r.label or '') for k in ('账面价值', '净值')):
            return r
    if rows:
        return rows[-1]
    details = [r for r in tm.rows if r.kind == 'detail' and r.label]
    return details[-1] if details else None


def _period_cols(tm):
    cols = {}
    for c in tm.columns:
        if c.is_data and c.period:
            cols.setdefault(c.period, c.index + 1)
    return cols


def _find_by_name(name, keys):
    """在一组名字里找最匹配的一个（互相包含，取最短的那个）。"""
    n = _norm(name)
    if not n:
        return None
    hits = [k for k in keys if _norm(k) and (_norm(k) == n)]
    if hits:
        return hits[0]
    hits = [k for k in keys if _norm(k) and (n in _norm(k) or _norm(k) in n)]
    if hits:
        return min(hits, key=lambda k: len(_norm(k)))
    return None


def match_all(book, model, number_index=None, item_map=None,
              amount_of=None):
    """返回 (matches, unmatched)。

    matches   : [Match]
    unmatched : [ReportItem]（报表有金额但没匹配到的，就是"漏项"）
    """
    by_sub = _tables_by_sub(model)
    sub_names = list(by_sub.keys())
    number_index = number_index or {}
    item_map = item_map or {}

    matches, unmatched, skipped = [], [], []
    for it in ri.ordered_items(book):
        if it.kind not in ('line', 'sub'):
            continue

        target = it.name
        how = '名称'
        if it.kind == 'sub' and it.parent:
            # 子项先落到父项目的标题上，再在表里找行
            target = it.parent

        sub_name = _find_by_name(target, sub_names)
        if sub_name is None and item_map.get(_norm(target)):
            sub_name = _find_by_name(item_map[_norm(target)], sub_names)
            how = '别名'
        if sub_name is None:
            want = number_index.get((it.sheet, it.row))
            if want and _find_by_name(want, sub_names):
                sub_name = _find_by_name(want, sub_names)
                how = '编号'
        if sub_name is None:
            # 报表是 0/空 → 正常没有这项，不算漏项（需求 7.1）
            if amount_of is not None and not amount_of(it):
                skipped.append(it)
            else:
                unmatched.append(it)
            continue

        tables = by_sub.get(sub_name) or []
        if not tables:
            unmatched.append(it)
            continue
        tm = tables[0]
        row = _pick_total_row(tm)
        if row is None:
            unmatched.append(it)
            continue
        matches.append(Match(it, sub_name, tm.ref.idx, row.row,
                             _period_cols(tm), how))
    return matches, unmatched, skipped


def summarize(matches, unmatched, skipped=()):
    from collections import Counter
    return {
        'matched': len(matches),
        'unmatched': len(unmatched),
        'no_money': len(skipped),
        'by_how': dict(Counter(m.how for m in matches)),
    }

def to_cross_checks(book, model, wb, existing, number_index=None,
                    tables=None):
    """把匹配结果转成 cross_checks.CrossCheck，供主流程直接使用。

    只补"现有勾稽没覆盖到、且报表有金额"的项目，不动已有结果。
    """
    import rules
    import cross_checks as cc

    covered = {(c.sheet, c.rows) for c in (existing or []) if c.sheet}
    # 用分析后的表模型（带 rows/columns）替换原始表引用
    if tables:
        by_sub = {}
        for tm in tables:
            by_sub.setdefault(tm.sub_item, []).append(tm)
        for ch in model.chapters:
            for sub in ch.sub_items:
                sub.tables = by_sub.get(sub.name, [])
    matches, _un, _sk = match_all(book, model, number_index=number_index)
    by_item = {id(m.item): m for m in matches}

    out = []
    for it in ri.ordered_items(book):
        if it.kind not in ('line', 'sub'):
            continue
        if (it.sheet, it.row) in covered:
            continue
        hit = by_item.get(id(it))
        if hit is None:
            continue
        sheet = next((s for s in book.sheets if s.name == it.sheet), None)
        if sheet is None:
            continue
        cols = sheet.blocks.get(it.block) or {}
        grid = _grid(model, hit.table_idx)
        if not grid:
            continue
        for col, period in cols.items():
            v = wb[it.sheet].cell(it.row, col).value
            if not isinstance(v, (int, float)) or abs(float(v)) < 0.005:
                continue
            ncol = hit.cols.get(period)
            if ncol is None:
                continue
            nv = rules.cell_number(grid, hit.row, ncol)
            if nv is None:
                continue
            out.append(cc.CrossCheck(
                chapter='', scope='', sub_item=hit.sub_item, name=it.name,
                period=period, period_key=None, note_value=nv, note_ref='',
                sheet=it.sheet, rows=it.row, cols=col,
                note_table=hit.table_idx, note_row=hit.row, note_col=ncol,
                note_terms=[(hit.row, ncol, 1)], basis='匹配层'))
    return out


def _grid(model, table_idx):
    for ch in model.chapters:
        for sub in ch.sub_items:
            for tm in sub.tables:
                if tm.ref.idx == table_idx:
                    return tm.ref.grid
    return []
