# -*- coding: utf-8 -*-
"""
rules.py — 勾稽规则库：从"看懂的表"生成一条条可执行的勾稽校验。

每条校验都能表达成「某个单元格 = 若干单元格的加减」，所以：
    · 在 Excel 里能写成活公式（点进去就看见引用了哪几格）
    · 在 Python 里也能直接算，立刻知道该条是平还是有差

规则分级
--------
hard（硬勾稽）：差异必须为 0，超过容差就是问题
soft（软提示）  ：比例重算这类，受四舍五入影响，只提示

容差
----
默认 0.01 元（1 分钱）。差额绝对值 ≤ 0.01 视为平。

已实现的规则
------------
V-SUM  合计行 = 其上方若干明细行之和（多合计行分别算；「其中：」子块不重复计入）
V-GRP  带「其中：」的父行 = 其下各明细之和
H-SUM  合计列 = 其余数据列之和
R-NET  账面价值 = 账面余额 − 坏账准备/减值准备（同一期间内）
R-NBV  账面价值 = 账面原值 − 累计折旧/摊销 − 减值准备（同一期间内）
"""

import datetime
import re
from dataclasses import dataclass, field
from typing import List, Optional

import docx_ir

TOLERANCE = 0.01
# 比例类容差：0.01 个百分点（不是 0.01 个百分点值本身）。
# 比例在表里存的是小数（53.61% = 0.5361），这里放大 100 倍后比较，
# 所以「四舍五入到两位小数」产生的尾差能过，再大就作为差异暴露出来，
# 由人工决定是否调整那一分钱尾差。
RATIO_TOLERANCE_PP = 0.01

# 视为"新的一段"的行标签：用于判断「其中：」子项到哪里结束
RE_BLOCK_BREAKER = re.compile(
    r'^(本期增加|本期减少|期初|期末|年初|年末|增加|减少'
    r'|账面原值|累计折旧|累计摊销|减值准备|账面价值|原值'
    r'|加：|减：|[①②③④⑤⑥⑦⑧⑨⑩])'
)
# 纯日期行（如「2026年9月30日」）同样是分段起点
RE_DATE_LABEL = re.compile(r'^\s*\d{4}\s*[年\-/.]\s*\d{1,2}')


def _is_breaker(label):
    label = label or ''
    return bool(RE_BLOCK_BREAKER.match(label) or RE_DATE_LABEL.match(label))


def row_sign(label):
    """行首写「减：」的，参与加总时取负号；「加：」取正号。"""
    label = (label or '').strip()
    if label.startswith('减：') or label.startswith('减:'):
        return -1
    return 1


# ---------------------------------------------------------------- 数据模型

@dataclass
class Term:
    coef: int          # +1 / -1
    row: int           # 表内 1 基行号
    col: int           # 表内 1 基列号
    table_idx: int = 0  # 0 = 本表；非 0 表示引用另一张表（跨表勾稽用）

    def key(self):
        return (self.table_idx, self.row, self.col)


@dataclass
class CheckItem:
    rule_id: str
    level: str                  # hard / soft
    chapter: str
    sub_item: str
    table_idx: int
    label: str                  # 展示文字
    target_row: int
    target_col: int
    terms: List[Term] = field(default_factory=list)   # 分子（或等式右边）
    den: List[Term] = field(default_factory=list)     # 分母；为空表示纯加减
    is_ratio: bool = False                            # 结果按百分比显示
    period: Optional[datetime.date] = None
    period_key: Optional[str] = None

    @property
    def table_uid(self):
        return f'{self.chapter}|{self.sub_item}|{self.table_idx}'

    def formula_refs(self):
        """给日志/调试看的表达式。"""
        parts = []
        for i, t in enumerate(self.terms):
            sign = '+' if t.coef > 0 else '−'
            if i == 0 and t.coef > 0:
                sign = ''
            parts.append(f'{sign}R{t.row}C{t.col}')
        expr = ''.join(parts)
        if self.den:
            dparts = []
            for i, t in enumerate(self.den):
                sign = '+' if t.coef > 0 else '−'
                if i == 0 and t.coef > 0:
                    sign = ''
                dparts.append(f'{sign}R{t.row}C{t.col}')
            expr = f'({expr})/({"".join(dparts)})'
        return expr


@dataclass
class CheckResult:
    item: CheckItem
    expected: Optional[float]
    actual: Optional[float]
    diff: Optional[float]
    status: str                 # pass / fail / skip(非数值)
    note: str = ''


# ---------------------------------------------------------------- 取值

def cell_number(grid, row, col):
    """取表内 1 基 (row, col) 的数值。'-' / 空 都当 0；非数值返回 None。"""
    r, c = row - 1, col - 1
    if r < 0 or r >= len(grid) or c < 0 or c >= len(grid[r]):
        return 0.0
    cell = grid[r][c]
    if cell.get('m') != ' ':
        return 0.0
    text = cell.get('t') or ''
    v = docx_ir.parse_cell_value(text)
    if v is None:
        return 0.0
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        return 0.0 if v.strip() in ('-', '--', '—', '——', '–', '/') else None
    return None


def evaluate(item, table_model, tables_by_idx=None):
    grid = table_model.ref.grid

    def _val(term):
        tid = term.table_idx or item.table_idx
        if tid == item.table_idx:
            return cell_number(grid, term.row, term.col)
        if tables_by_idx and tid in tables_by_idx:
            other = tables_by_idx[tid]
            return cell_number(other.ref.grid, term.row, term.col)
        return None

    total = 0.0
    for t in item.terms:
        v = _val(t)
        if v is None:
            return CheckResult(item, None, None, None, 'skip', '运算项里有非数字或找不到表')
        total += t.coef * v

    expected = total
    if item.den:
        den = 0.0
        for t in item.den:
            v = _val(t)
            if v is None:
                return CheckResult(item, None, None, None, 'skip', '分母里有非数字或找不到表')
            den += t.coef * v
        if abs(den) < 1e-9:
            return CheckResult(item, None, None, None, 'skip', '分母为 0，比例无法计算')
        expected = total / den

    actual = cell_number(grid, item.target_row, item.target_col)
    if actual is None:
        return CheckResult(item, expected, None, None, 'skip', '目标格不是数字')
    diff = actual - expected
    if item.is_ratio:
        # 比例按"百分点"比较：0.01 个百分点以内算平
        status = 'pass' if abs(diff) * 100 <= RATIO_TOLERANCE_PP else 'fail'
        # 比例类一律按"软提示"处理
        if item.level == 'hard':
            item.level = 'soft'
    else:
        status = 'pass' if abs(diff) <= TOLERANCE else 'fail'
    return CheckResult(item, expected, actual, diff, status)


# ---------------------------------------------------------------- 行/列工具

def _data_columns(tm, include_ratio=False):
    out = []
    for c in tm.columns:
        if not c.is_data:
            continue
        if c.kind == 'ratio' and not include_ratio:
            continue
        if c.kind == 'text':
            continue
        out.append(c)
    return out


def _consume_children(tm, block, start, target_value, ref_col, tolerance=TOLERANCE):
    """从 block[start] 起，吃掉属于上层的子项行，返回消耗到的位置。

    判定依据：子项金额累加到与目标值相等时，这一块结束。
    这是"先用一列认出结构，再逐列独立校验"——结构识别和结果校验分开，
    不会因为挑行而让错误消失。
    """
    grid = tm.ref.grid
    rows = {r.row: r for r in tm.rows}
    acc = 0.0
    i = start
    while i < len(block):
        info = rows[block[i]]
        if info.kind != 'detail' or _is_breaker(info.label):
            break
        v = cell_number(grid, block[i], ref_col)
        acc += abs(v or 0.0)
        i += 1
        if target_value is not None and abs(acc - abs(target_value)) <= tolerance:
            break
    return i


# 这些行是"计算基础"/说明行，不参与上方加总
STOP_LABELS = ('利润总额', '营业收入', '营业成本', '利润总额（亏损总额）')


def _pick_ref_col(tm, row, cols):
    """挑一个"有值"的列作为结构识别参考列。

    不能固定用第一列数据列：应付职工薪酬表的第 1 列是期初数，多半是「-」，
    用它判断"子项累加到多少"会失准。
    """
    for c in cols:
        v = cell_number(tm.ref.grid, row, c.index + 1)
        if v:
            return c.index + 1
    return cols[0].index + 1 if cols else 1


def _contribution_rows(tm, total_row, ref_col):
    """求某个合计行上方的"代表行"——这些行的和应当等于合计。

    层级不能重复计：
      · 「主营业务小计」：它自己代表下面的明细 → 取小计，跳过其明细
      · 「其中：购置」：它是上一行的拆分 → 跳过它和它的子项
      · 子项到哪里结束：用金额匹配判定（累加到等于上层金额为止）
      · 「利润总额」这类计算基础行：不参与加总，直接截断
    """
    rows = {r.row: r for r in tm.rows}

    block = []
    r = total_row - 1
    while r >= 1:
        info = rows.get(r)
        if info is None:
            break
        if info.kind in ('blank', 'section', 'header', 'title', 'total'):
            break
        if (info.label or '').strip() in STOP_LABELS:
            break
        block.append(r)
        r -= 1
    block.sort()

    picked = []
    i = 0
    while i < len(block):
        row = block[i]
        info = rows[row]

        if info.kind == 'subtotal':
            picked.append(row)
            val = cell_number(tm.ref.grid, row, ref_col)
            nxt = _consume_children(tm, block, i + 1, val, ref_col)
            i = nxt if nxt > i else i + 1
            continue

        if info.kind == 'group':
            parent_val = cell_number(tm.ref.grid, picked[-1], ref_col) if picked else None
            nxt = _consume_children(tm, block, i + 1, parent_val, ref_col)
            i = nxt if nxt > i else i + 1
            continue

        if info.kind == 'detail':
            if _is_breaker(info.label):
                picked.append(row)
                i += 1
                continue
            picked.append(row)
            i += 1
            continue

        i += 1
    return picked


def _group_children(tm, group_row, parent_row, limit_row, ref_col, tolerance=TOLERANCE):
    """取「其中：」这一组的明细行：包含 group 行自身（它常自带金额），
    再加上它下面直到**累加金额等于父行金额**为止的行。"""
    grid = tm.ref.grid
    rows = {r.row: r for r in tm.rows}
    parent_val = cell_number(grid, parent_row, ref_col)

    out = []
    acc = 0.0
    r = group_row
    while r <= limit_row:
        info = rows.get(r)
        if info is None:
            break
        if info.kind not in ('detail', 'group'):
            break
        if r != group_row and _is_breaker(info.label):
            break
        out.append(r)
        acc += abs(cell_number(grid, r, ref_col) or 0.0)
        r += 1
        if parent_val is not None and abs(acc - abs(parent_val)) <= tolerance:
            break
    return out


def _cols_by_role(tm, period_key=None, period=None):
    """按 (期间, 角色) 归集列号（1 基）。"""
    out = {}
    for c in tm.columns:
        if not c.is_data or not c.role:
            continue
        if period_key is not None and c.period_key != period_key:
            continue
        if period is not None and c.period != period:
            continue
        out.setdefault(c.role, []).append(c)
    return out


# ---------------------------------------------------------------- 规则实现

def _rule_v_sum(tm, checks):
    grid = tm.ref.grid
    totals = [r for r in tm.rows if r.kind == 'total']
    if not totals:
        return
    cols = _data_columns(tm, include_ratio=False)
    rows = {r.row: r for r in tm.rows}
    for tot in totals:
        ref_col = _pick_ref_col(tm, tot.row, cols)
        detail_rows = _contribution_rows(tm, tot.row, ref_col)
        if not detail_rows:
            continue
        for c in cols:
            terms = [Term(row_sign(rows[r].label), r, c.index + 1) for r in detail_rows]
            checks.append(CheckItem(
                rule_id='V-SUM', level='hard',
                chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
                label=f'合计行 = 上方 {len(detail_rows)} 行之和（{c.header_text or "列" + str(c.index + 1)}）',
                target_row=tot.row, target_col=c.index + 1,
                terms=terms, period=c.period, period_key=c.period_key,
            ))


def _rule_v_grp(tm, checks):
    """带「其中：」的父行 = 其下明细之和。"""
    groups = [r for r in tm.rows if r.kind == 'group']
    if not groups:
        return
    rows = {r.row: r for r in tm.rows}
    cols = _data_columns(tm)
    for g in groups:
        parent_row = g.row - 1
        parent = rows.get(parent_row)
        if parent is None or parent.kind != 'detail':
            continue
        ref_col = _pick_ref_col(tm, parent_row, cols)
        children = _group_children(tm, g.row, parent_row, len(tm.rows), ref_col)
        if not children:
            continue
        for c in cols:
            terms = [Term(row_sign(rows[r].label), r, c.index + 1) for r in children]
            checks.append(CheckItem(
                rule_id='V-GRP', level='hard',
                chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
                label=f'「{parent.label}」= 其下 {len(children)} 个明细之和',
                target_row=parent_row, target_col=c.index + 1,
                terms=terms, period=c.period, period_key=c.period_key,
            ))


def _rule_h_sum(tm, checks):
    """列向加减：合计列 = 其余数据列之和。

    分两种情况，避免重复计算：
      · 表里有「小计」列 → 合计列 = 各小计列之和（不再把明细列再加一遍）
      · 没有小计列      → 合计列 = 其余明细列之和

    只认「最末一级表头就是 合计/总计/小计」的列。不能只看表头里含"合计"——
    像「占应收账款合计的比例」这种列也含"合计"，但它是比例列，不是加总列。
    """
    total_words = ('合计', '总计')
    sub_words = ('小计',)

    def leaf(c):
        return (c.path[-1].strip() if c.path else '')

    total_cols = [c for c in tm.columns if c.is_data and leaf(c) in total_words]
    if not total_cols:
        return
    sub_cols = [c for c in tm.columns if c.is_data and leaf(c) in sub_words]

    if sub_cols:
        terms_cols = [c for c in sub_cols if c not in total_cols]
        label = f'合计列 = {len(terms_cols)} 个小计列之和'
    else:
        terms_cols = [c for c in _data_columns(tm)
                      if c not in total_cols
                      and leaf(c) not in total_words
                      and leaf(c) not in sub_words]
        label = f'合计列 = 其余 {len(terms_cols)} 列之和'

    if len(terms_cols) < 1:
        return

    for r in tm.rows:
        if r.kind not in ('detail', 'total', 'subtotal'):
            continue
        if not r.label:
            continue
        for tc in total_cols:
            terms = [Term(+1, r.row, c.index + 1) for c in terms_cols]
            checks.append(CheckItem(
                rule_id='H-SUM', level='hard',
                chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
                label=f'{label}（{r.label}）',
                target_row=r.row, target_col=tc.index + 1,
                terms=terms,
            ))


def _rule_balance(tm, checks):
    """同一期间内：价值 = 余额 − 准备；价值 = 原值 − 折旧 − 减值。"""
    # 按 (期初/期末/日期) 分组
    groups = {}
    for c in tm.columns:
        if not c.is_data or not c.role:
            continue
        key = c.period if c.period else c.period_key
        if key is None:
            continue
        groups.setdefault(key, []).append(c)

    for key, cols in groups.items():
        roles = {}
        for c in cols:
            roles.setdefault(c.role, []).append(c)

        rows = [r for r in tm.rows if r.kind in ('detail', 'total') and r.label]

        # 价值 = 余额 − 准备
        if '价值' in roles and '余额' in roles and '准备' in roles:
            for r in rows:
                terms = [Term(+1, r.row, roles['余额'][0].index + 1),
                         Term(-1, r.row, roles['准备'][0].index + 1)]
                checks.append(CheckItem(
                    rule_id='R-NET', level='hard',
                    chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
                    label=f'账面价值 = 账面余额 − 坏账准备（{r.label}）',
                    target_row=r.row, target_col=roles['价值'][0].index + 1,
                    terms=terms,
                ))

        # 价值 = 原值 − 折旧 − 减值
        if '价值' in roles and '原值' in roles and ('折旧' in roles or '准备' in roles):
            for r in rows:
                terms = [Term(+1, r.row, roles['原值'][0].index + 1)]
                if '折旧' in roles:
                    terms.append(Term(-1, r.row, roles['折旧'][0].index + 1))
                if '准备' in roles:
                    terms.append(Term(-1, r.row, roles['准备'][0].index + 1))
                checks.append(CheckItem(
                    rule_id='R-NBV', level='hard',
                    chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
                    label=f'账面价值 = 账面原值 − 累计折旧 − 减值（{r.label}）',
                    target_row=r.row, target_col=roles['价值'][0].index + 1,
                    terms=terms,
                ))


def _rule_move(tm, checks):
    """期末 = 期初 + 本期增加 − 本期减少。

    两种排版都要支持：
      · 行式（固定资产）：块内依次是 期初行 / 本期增加金额 / 本期减少金额 / 期末行
      · 列式（无形资产）：表头依次是 期初 / 本期增加 / 本期减少 / 期末
    """
    rows = {r.row: r for r in tm.rows}
    # 这里不能只用"有数字"的数据列：某期可能整列为空（如本期减少），
    # 但仍要参与 期末 = 期初 + 增加 − 减少 的校验。
    cols = [c for c in tm.columns if c.is_data]

    # ---------- 列式 ----------
    inc_cols = [c for c in cols if c.role == '增加']
    dec_cols = [c for c in cols if c.role == '减少']
    if inc_cols and dec_cols:
        inc, dec = inc_cols[0].index, dec_cols[0].index
        beg_col = next((c for c in cols if c.index == inc - 1), None)
        end_col = next((c for c in cols if c.index == dec + 1), None)
        if beg_col and end_col and end_col.role != '增加' and end_col.role != '减少':
            for r in tm.rows:
                if r.kind not in ('detail', 'total') or not r.label:
                    continue
                # 账面价值这类是倒算出来的，不服从"期初+增加−减少"
                if any(k in (r.label or '') for k in ('账面价值', '净值', '净额')):
                    continue
                checks.append(CheckItem(
                    rule_id='R-MOVE', level='hard',
                    chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
                    label=f'期末 = 期初 + 本期增加 − 本期减少（{r.label}）',
                    target_row=r.row, target_col=end_col.index + 1,
                    terms=[Term(+1, r.row, beg_col.index + 1),
                           Term(+1, r.row, inc_cols[0].index + 1),
                           Term(-1, r.row, dec_cols[0].index + 1)],
                ))

    # ---------- 行式 ----------
    inc_rows = [r for r in tm.rows if r.kind == 'detail'
                and (r.label or '').startswith('本期增加')]
    dec_rows = [r for r in tm.rows if r.kind == 'detail'
                and (r.label or '').startswith('本期减少')]
    for inc in inc_rows:
        # 期初 = 本期增加上面最近的一行明细
        beg = None
        for r in range(inc.row - 1, 0, -1):
            info = rows.get(r)
            if info is None or info.kind in ('blank', 'section', 'header', 'title', 'total'):
                break
            beg = info
            break
        dec = next((d for d in dec_rows if d.row > inc.row), None)
        if beg is None or dec is None:
            continue
        # 期末 = 本期减少之后第一行"日期"或"期末余额"行
        end = None
        for r in range(dec.row + 1, len(tm.rows) + 1):
            info = rows.get(r)
            if info is None:
                break
            if info.kind in ('blank', 'section', 'header', 'title', 'total'):
                break
            label = (info.label or '').strip()
            if RE_DATE_LABEL.match(label) or label.startswith(('期末', '本期末', '年末')):
                end = info
                break
        if end is None:
            continue
        for c in cols:
            checks.append(CheckItem(
                rule_id='R-MOVE', level='hard',
                chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
                label=f'期末 = 期初 + 本期增加 − 本期减少（{end.label} / {c.header_text or "列"}）',
                target_row=end.row, target_col=c.index + 1,
                terms=[Term(+1, beg.row, c.index + 1),
                       Term(+1, inc.row, c.index + 1),
                       Term(-1, dec.row, c.index + 1)],
                period=c.period, period_key=c.period_key,
            ))


def _rule_pct(tm, checks, siblings=None):
    """比例勾稽（软提示，四舍五入必然有尾差）。

    三种情形：
      · 列名带「计提」→ 计提比例 = 坏账准备 ÷ 账面余额
      · 前五名/关联方表 → 占比 = 该行金额 ÷ **该项目总额**（取同子项目主表合计，跨表引用）
      · 其余比例列    → 占比 = 该行金额 ÷ 本表合计行金额
    """
    cols = [c for c in tm.columns if c.is_data]
    ratio_cols = [c for c in cols if c.kind == 'ratio' or c.role == '比例']
    if not ratio_cols:
        return

    def col_of(period, roles):
        for c in cols:
            if c.period != period:
                continue
            if c.role in roles:
                return c
        return None

    def base_table(period):
        """同子项目里能找到"该项目总额"的表：有合计行、有金额列。

        只认出现在本表**之前**的表——成果表是按文档顺序写的，
        后面的表此刻还没定行号，没法写跨表公式。
        前五名表通常表头不写期间，这时取主表**最近一期**的合计。
        """
        for other in (siblings or []):
            if other is tm or other.ref.idx >= tm.ref.idx:
                continue
            tot = next((r for r in other.rows if r.kind == 'total'), None)
            if tot is None:
                continue
            cols = [c for c in other.columns
                    if c.is_data and c.role in ('余额', '金额', '价值')]
            if not cols:
                continue
            if period is not None:
                col = next((c for c in cols if c.period == period), None)
            else:
                dated = [c for c in cols if c.period]
                latest = max((c.period for c in dated), default=None)
                col = next((c for c in dated if c.period == latest), None)
            if col is None:
                continue
            return other, tot, col
        return None

    total_rows = [r for r in tm.rows if r.kind == 'total']
    tot = total_rows[-1] if total_rows else None
    targets = [r for r in tm.rows if r.kind in ('detail', 'subtotal') and r.label]
    is_top5 = getattr(tm, 'table_type', '') == 'TOP5'

    for rc in ratio_cols:
        period = rc.period
        head = rc.header_text or ''
        if '计提' in head:
            prov = col_of(period, ('准备',))
            base = col_of(period, ('余额', '金额'))
            if not (prov and base):
                continue
            for r in targets:
                checks.append(CheckItem(
                    rule_id='R-PCT', level='soft',
                    chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
                    label=f'计提比例 = 坏账准备 ÷ 账面余额（{r.label}）',
                    target_row=r.row, target_col=rc.index + 1,
                    terms=[Term(+1, r.row, prov.index + 1)],
                    den=[Term(+1, r.row, base.index + 1)],
                    is_ratio=True, period=period, period_key=rc.period_key,
                ))
            continue

        base = col_of(period, ('余额', '金额', '价值'))
        if base is None:
            continue

        if is_top5:
            # 前五名表的占比分母是"该项目总额"，从同子项目主表的合计行取
            found = base_table(period)
            if found is None:
                continue
            other, other_tot, other_col = found
            targets_incl_total = [r for r in tm.rows
                                  if r.kind in ('detail', 'subtotal', 'total') and r.label]
            for r in targets_incl_total:
                checks.append(CheckItem(
                    rule_id='R-PCT', level='soft',
                    chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
                    label=f'占比 = 该行金额 ÷ {other.sub_item}合计'
                          f'（{r.label}，分母取自 表#{other.ref.idx}）',
                    target_row=r.row, target_col=rc.index + 1,
                    terms=[Term(+1, r.row, base.index + 1)],
                    den=[Term(+1, other_tot.row, other_col.index + 1,
                              table_idx=other.ref.idx)],
                    is_ratio=True, period=period, period_key=rc.period_key,
                ))
            continue

        if tot is None:
            continue
        for r in targets:
            checks.append(CheckItem(
                rule_id='R-PCT', level='soft',
                chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
                label=f'占比 = 该行金额 ÷ 合计（{r.label}）',
                target_row=r.row, target_col=rc.index + 1,
                terms=[Term(+1, r.row, base.index + 1)],
                den=[Term(+1, tot.row, base.index + 1)],
                is_ratio=True, period=period, period_key=rc.period_key,
            ))


# ---------------------------------------------------------------- 入口

def build_checks(tables):
    """对全部表生成勾稽校验。"""
    checks = []
    by_sub = {}
    for tm in tables:
        by_sub.setdefault((tm.chapter, tm.sub_item), []).append(tm)

    for tm in tables:
        siblings = by_sub.get((tm.chapter, tm.sub_item), [])
        _rule_v_sum(tm, checks)
        _rule_v_grp(tm, checks)
        _rule_h_sum(tm, checks)
        _rule_balance(tm, checks)
        _rule_move(tm, checks)
        _rule_pct(tm, checks, siblings)
    return checks


def evaluate_all(checks, tables_by_uid):
    """逐条算，返回 CheckResult 列表。"""
    idx_map = {tm.ref.idx: tm for tm in tables_by_uid.values()}
    results = []
    for item in checks:
        tm = tables_by_uid.get(item.table_uid)
        if tm is None:
            results.append(CheckResult(item, None, None, None, 'skip', '找不到表'))
            continue
        results.append(evaluate(item, tm, idx_map))
    return results


def map_tables(tables):
    return {f'{t.chapter}|{t.sub_item}|{t.ref.idx}': t for t in tables}


def summarize(results):
    from collections import Counter
    c = Counter(r.status for r in results)
    return dict(c)
