# -*- coding: utf-8 -*-
"""
cross_checks.py — 附注合计数 与 三张报表科目数 的跨表勾稽。

做法
----
1. 每个子项目取其**第一张有"合计"行且列期间能定位**的表，读合计行的各期金额
2. 用子项目名去报表里按科目名找行（合并/母公司分别对应）
3. 按期间日期找报表的列
4. 差额 = 附注合计 − 报表数

报表找不到科目、或期间对应不上，不算差异，记为"未匹配"（待确认）。
"""

import datetime
import re
from dataclasses import dataclass
from typing import List, Optional

import report_reader as rr
import rules


@dataclass
class CrossCheck:
    chapter: str
    scope: str
    sub_item: str
    name: str
    period: Optional[datetime.date]
    period_key: Optional[str]
    note_value: Optional[float]
    note_ref: str                 # 如 "表#17 R4 C2"
    note_table: int = 0           # 附注表号（用于写到底稿右侧）
    note_row: int = 0             # 附注表内行号
    note_col: int = 0             # 附注表内列号
    note_terms: list = None       # [(行, 列, 系数)]：口径为"余额−准备"时要两项
    sheet: str = ''
    rows: Optional[int] = None
    cols: Optional[int] = None
    report_value: Optional[float] = None
    diff: Optional[float] = None
    status: str = 'unmatched'     # pass / fail / unmatched
    note: str = ''
    basis: str = ''               # 取数口径：账面价值 / 账面余额−坏账准备 / 账面余额


def _total_row_of(tm):
    for r in tm.rows:
        if r.kind == 'total':
            return r
    return None


def _pick_main_table(tables, names=None):
    """挑一张最能代表该子项目金额的表。

    附注的排版习惯是：**子项目下的第一张表就是主披露表**（后面是分项、
    变动明细等）。所以按文档顺序取第一张合格的，而不是按列数打分——
    否则应收账款会拿到「组合计提」那种分表，金额口径就错了。
    """
    # 第一轮：有「合计」行且期间可定位
    for tm in tables:
        tot = _total_row_of(tm)
        if tot is None:
            continue
        if any(c.period for c in tm.columns if c.is_data):
            return tm, tot
    # 第二轮：没有合计行，但有「账面价值」/「期末xx」行或只有一个明细行
    for tm in tables:
        tot = _summary_row_of(tm, names)
        if tot is None:
            continue
        if any(c.period for c in tm.columns if c.is_data):
            return tm, tot
    return None, None


SUMMARY_LABELS = ('账面价值', '期末账面价值', '净额', '净值')

# 子项目名与报表科目名对不上时的别名
ALIAS_NAMES = {
    '现金流量表补充资料': ['经营活动产生的现金流量净额', '现金及现金等价物净增加额'],
    '营业收入和营业成本': ['营业收入', '营业成本'],
}


def _summary_row_of(tm, names=None):
    """代表该子项目对外披露金额的行。

    依次尝试：
      1. 「合计」行
      2. 「账面价值」行（无形资产这类没有合计）
      3. 标签里同时含子项目名与「期末/年末」的行（如「期末未分配利润」）
      4. 表里**只有一个**明细行时，就用它（合同负债这类只披露一行）
    """
    tot = _total_row_of(tm)
    if tot is not None:
        return tot
    for r in tm.rows:
        if r.kind == 'detail' and (r.label or '').strip() in SUMMARY_LABELS:
            return r
    if names:
        # 标签与子项目名（或别名）完全一致，或长名字被包含
        for r in tm.rows:
            if r.kind != 'detail':
                continue
            label = (r.label or '').replace(' ', '')
            for n in names:
                if not n:
                    continue
                if label == n or (len(n) >= 6 and n in label):
                    return r
        for r in tm.rows:
            if r.kind != 'detail':
                continue
            label = (r.label or '').replace(' ', '')
            # 必须以「期末/年末」开头：否则会选到「调整前上期末…」这种上期数
            if label.startswith(('期末', '年末')) and \
                    any(n and n in label for n in names):
                return r
    details = [r for r in tm.rows if r.kind == 'detail' and r.label]
    if len(details) == 1:
        return details[0]
    return None


def _split_names(name):
    """把子项目名拆成能在报表里单独匹配的名字。

    含"和/、/及"时**不要**再拿全名去匹配——全名会被包含匹配误命中
    "营业收入"，导致取到成本列。直接按拆分名逐个匹配。
    """
    if name in ALIAS_NAMES:
        return ALIAS_NAMES[name]
    parts = [p.strip() for p in re.split(r'[和、及与]', name) if len(p.strip()) >= 2]
    if parts:
        return parts
    return [name]


def _match_report(report, scope, names):
    """在报表里按名字找科目，返回 (sheet, (row, col), 命中的名字)。"""
    for cand_name in names:
        for kind in ('资产负债表', '利润表', '现金流量表'):
            sheet = report.find(scope, kind)
            if sheet is None and scope != '单体':
                sheet = report.find('单体', kind)
            if sheet is None:
                continue
            pos = rr.find_item(sheet, cand_name)
            if pos:
                return sheet, pos, cand_name
    return None, None, None


def build_cross_checks(model, tables, report):
    by_sub = {}
    for tm in tables:
        by_sub.setdefault((tm.chapter, tm.sub_item), []).append(tm)

    out = []
    for ch in model.chapters:
        for sub in ch.sub_items:
            group = by_sub.get((ch.name, sub.name), [])
            names = _split_names(sub.name)
            tm, tot = _pick_main_table(group, names)
            if tm is None:
                continue

            scope = ch.scope
            sheet, found, matched_name = _match_report(report, scope, names)
            if sheet is None or found is None:
                out.append(CrossCheck(
                    chapter=ch.name, scope=scope, sub_item=sub.name, name=sub.name,
                    period=None, period_key=None, note_value=None, note_ref='',
                    status='unmatched', note='报表里没找到同名科目'))
                continue
            item_row, item_col = found

            # 命中"营业收入"这类拆分名时，只取表头能对上的列。
            # 表头常写「收入」「成本」而不是全称，所以再试名字的后两个字。
            hint = matched_name if matched_name != sub.name else None
            hint_keys = []
            if hint:
                hint_keys = [hint]
                if len(hint) >= 2:
                    hint_keys.append(hint[-2:])

            def _col_ok(col):
                if not hint_keys:
                    return True
                head = (col.header_text or '').replace(' ', '')
                return any(k in head for k in hint_keys)

            per_period = {}
            for col in tm.columns:
                if not col.is_data or not col.period:
                    continue
                if not _col_ok(col):
                    continue
                role = col.role or ('金额' if col.kind == 'amount' else None)
                if role in ('价值', '余额', '金额', '准备'):
                    # 同一期间同一角色有多列时（如"收入/成本"都判成金额），
                    # 保留文档顺序里的第一列，别让后面的覆盖前面的
                    per_period.setdefault(col.period, {}).setdefault(role, col)
            if hint and not per_period:
                # 表头完全对不上，退回用全部列，但标记出来
                for col in tm.columns:
                    if not col.is_data or not col.period:
                        continue
                    role = col.role or ('金额' if col.kind == 'amount' else None)
                    if role in ('价值', '余额', '金额', '准备'):
                        per_period.setdefault(col.period, {})[role] = col

            for period, roles in sorted(per_period.items()):
                if '价值' in roles:
                    col, basis = roles['价值'], '账面价值'
                    note_val = rules.cell_number(tm.ref.grid, tot.row, col.index + 1)
                    terms = [(tot.row, col.index + 1, 1)]
                elif '余额' in roles and '准备' in roles:
                    col, basis = roles['余额'], '账面余额−坏账准备'
                    bal = rules.cell_number(tm.ref.grid, tot.row, roles['余额'].index + 1)
                    prov = rules.cell_number(tm.ref.grid, tot.row, roles['准备'].index + 1)
                    note_val = None if (bal is None or prov is None) else bal - prov
                    terms = [(tot.row, roles['余额'].index + 1, 1),
                             (tot.row, roles['准备'].index + 1, -1)]
                elif '余额' in roles:
                    col, basis = roles['余额'], '账面余额'
                    note_val = rules.cell_number(tm.ref.grid, tot.row, col.index + 1)
                    terms = [(tot.row, col.index + 1, 1)]
                elif '金额' in roles:
                    col, basis = roles['金额'], '金额'
                    note_val = rules.cell_number(tm.ref.grid, tot.row, col.index + 1)
                    terms = [(tot.row, col.index + 1, 1)]
                else:
                    continue
                if note_val is None:
                    continue

                cc = CrossCheck(
                    chapter=ch.name, scope=scope, sub_item=sub.name,
                    name=sub.name, period=col.period, period_key=col.period_key,
                    note_value=note_val, basis=basis,
                    note_ref=f'表#{tm.ref.idx} R{tot.row} C{col.index + 1}',
                    note_table=tm.ref.idx, note_row=tot.row, note_col=col.index + 1,
                    note_terms=terms,
                )
                col_no = sheet.value_col(item_col, col.period)
                if col_no is None:
                    cc.sheet = sheet.name
                    cc.rows = item_row
                    cc.cols = item_col
                    cc.status = 'unmatched'
                    cc.note = f'报表里没有 {col.period} 这一期'
                    out.append(cc)
                    continue

                cc.sheet = sheet.name
                cc.rows = item_row
                cc.cols = col_no
                out.append(cc)
    return out


NIL_VALUES = {'-', '--', '—', '——', '–', '/', '', '无', '不适用'}


def to_number(value):
    """把报表单元格转成数字。'-' 等空值标记按 0 处理，无法解析返回 None。"""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if value is None:
        return 0.0
    s = str(value).strip().replace(',', '')
    if s in NIL_VALUES:
        return 0.0
    if s.endswith('%'):
        try:
            return float(s[:-1]) / 100.0
        except ValueError:
            return None
    try:
        return float(s)
    except ValueError:
        return None


def evaluate(cross, report_values):
    """report_values: {(sheet_name,row,col): value}"""
    if cross.sheet and cross.rows and cross.cols:
        raw = report_values.get((cross.sheet, cross.rows, cross.cols))
        cross.report_value = to_number(raw)
        if cross.report_value is None:
            cross.status = 'unmatched'
            cross.note = f'报表取数不是数字：{raw!r}'
            return cross
    else:
        cross.report_value = None
    if cross.report_value is None:
        if cross.status != 'unmatched':
            cross.status = 'unmatched'
            cross.note = cross.note or '报表取数为空'
        return cross
    cross.diff = cross.note_value - cross.report_value
    cross.status = 'pass' if abs(cross.diff) <= rules.TOLERANCE else 'fail'
    return cross
