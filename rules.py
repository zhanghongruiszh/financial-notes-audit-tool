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
    """从 block[start] 起，试着吃掉属于上层的子项行。

    返回 (结束位置, 是否**正好**累加到目标值)。
    只有"正好对上"才算真的子项——否则像「其中：开发成本」这种
    本身就是上一行全额重述的行，会把后面互不相干的行（库存商品…）全吞掉。
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
            return i, True
        if target_value is not None and acc > abs(target_value) + tolerance:
            break          # 再加就超了，说明这些不是它的子项
    return i, False


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
        if info.kind in ('blank', 'section', 'header', 'title'):
            break
        if (info.label or '').strip() in STOP_LABELS:
            break
        if info.kind == 'total':
            # 上方遇到「小计」：
            #   · 目标是「合计」→ 它是组成项（合计 = 小计 − 减值准备），收进来
            #   · 目标是「小计」→ 上面的合计是更高层级，不属于它，停
            target_label = (rows.get(total_row).label or '') if rows.get(total_row) else ''
            if '合计' in target_label or '总计' in target_label:
                block.append(r)
            break
        block.append(r)
        r -= 1
    block.sort()

    picked = []
    i = 0
    while i < len(block):
        row = block[i]
        info = rows[row]

        if info.kind in ('total', 'subtotal'):
            picked.append(row)
            if info.kind == 'total':
                i += 1          # 小计不再向下展开（下面已由它代表）
                continue
            val = cell_number(tm.ref.grid, row, ref_col)
            nxt, matched = _consume_children(tm, block, i + 1, val, ref_col)
            i = nxt if (matched and nxt > i) else i + 1
            continue

        if info.kind == 'group':
            parent_val = cell_number(tm.ref.grid, picked[-1], ref_col) if picked else None
            nxt, matched = _consume_children(tm, block, i + 1, parent_val, ref_col)
            i = nxt if (matched and nxt > i) else i + 1
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
    """合计行 = 上方明细之和（含小计/其中/减项处理）。"""
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

    # 四块式变动表（固定资产/无形资产）由 R-BLK 负责，列式规则不适用
    if _find_blocks(tm):
        return

    # ---------- 列式 ----------
    inc_cols = [c for c in cols if c.role == '增加']
    dec_cols = [c for c in cols if c.role == '减少']
    if inc_cols and dec_cols:
        inc, dec = inc_cols[0], dec_cols[0]

        def head_of(c):
            return c.header_text or ''

        # 期初/期末按表头认，不能靠位置——中间可能夹着「本年重分类」这类列
        beg_col = next((c for c in cols if c.role == '余额'
                        and any(k in head_of(c) for k in ('期初', '年初'))), None)
        end_col = next((c for c in cols if c.role == '余额'
                        and any(k in head_of(c) for k in ('期末', '年末'))), None)
        if beg_col is None:
            beg_col = next((c for c in cols if c.index == inc.index - 1), None)
        if end_col is None:
            end_col = next((c for c in cols if c.index == dec.index + 1), None)
        if beg_col and end_col and end_col is not dec_cols[0] and end_col is not inc_cols[0]:
            # 增加/减少之外的其他变动列（重分类、其他）也算进等式
            extras = [c for c in cols
                      if dec.index < c.index < end_col.index
                      and c.role not in ('增加', '减少', '余额', '价值', '准备', '比例')]
            for r in tm.rows:
                if r.kind not in ('detail', 'total', 'subtotal') or not r.label:
                    continue
                # 账面价值这类是倒算出来的，不服从"期初+增加−减少"
                if any(k in (r.label or '') for k in ('账面价值', '净值', '净额')):
                    continue
                terms = [Term(+1, r.row, beg_col.index + 1),
                         Term(+1, r.row, inc.index + 1),
                         Term(-1, r.row, dec.index + 1)]
                terms += [Term(+1, r.row, c.index + 1) for c in extras]
                checks.append(CheckItem(
                    rule_id='R-MOVE', level='hard',
                    chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
                    label=f'期末 = 期初 + 本期增加 − 本期减少（{r.label}）',
                    target_row=r.row, target_col=end_col.index + 1,
                    terms=terms,
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
        """同子项目里找"该项目总额"：有合计行、有金额列。

        只认排在本表**之前**的表——成果表按文档顺序写，后面的表还没定行号。
        候选可能不止一张（分类表、账龄表、三阶段表…），取**合计最大**的那张：
        小口径的附表（某类资产的账龄）合计会偏小，拿它当分母会把占比算大。
        """
        best = None
        for other in (siblings or []):
            if other is tm or other.ref.idx >= tm.ref.idx:
                continue
            tot = next((r for r in other.rows if r.kind in ('total', 'subtotal')), None)
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
            val = cell_number(other.ref.grid, tot.row, col.index + 1)
            if val is None:
                continue
            if best is None or abs(val) > abs(best[3]):
                best = (other, tot, col, val)
        return best[:3] if best else None

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

def _rule_sum_down(tm, checks):
    """合计在**上方**、明细/小计在下方：合计 = 各小计之和。

    长期股权投资那种表就是：第一行「合计」，下面分「子公司…小计」「联营企业…小计」。
    """
    rows = [r for r in tm.rows if r.row]
    if not rows:
        return
    first_total = next((r for r in rows if r.kind == 'total'), None)
    if first_total is None:
        return
    # 只处理"合计出现在表头附近、上面没有别的合计"的版式
    above = [r for r in rows if r.row < first_total.row
             and r.kind in ('detail', 'group', 'subtotal', 'total')]
    if above:
        return

    subs = [r for r in rows if r.row > first_total.row
            and r.kind in ('subtotal', 'total')]
    if not subs:
        # 合计在上、下面直接是明细（没有小计）：合计 = 全部明细之和
        subs = [r for r in rows if r.row > first_total.row
                and r.kind == 'detail'
                and not (r.label or '').startswith(('其中', '减：', '减:', '加：', '--'))]
    if not subs:
        return

    for c in _data_columns(tm):
        terms = [Term(+1, r.row, c.index + 1) for r in subs]
        checks.append(CheckItem(
            rule_id='R-SUMD', level='hard',
            chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
            label=f'合计 = 下方 {len(subs)} 行之和（{c.header_text or "列"}）',
            target_row=first_total.row, target_col=c.index + 1,
            terms=terms,
        ))


# ---------------------------------------------------------------- 特殊表结构

# 固定资产/无形资产这类"分块变动表"的块标题
BLOCK_KEYWORDS = (
    ('原值', ('账面原值', '原值')),
    ('折旧', ('累计折旧', '累计摊销', '折旧')),
    ('减值', ('减值准备', '减值')),
    ('价值', ('账面价值',)),
)


def _find_blocks(tm):
    """把行按「① 账面原值 / ② 累计折旧 / ③ 减值准备 / ④ 账面价值」分块。

    返回 [(块类型, [RowInfo...]), ...]，块内不含块标题行。
    """
    rows = [r for r in tm.rows if r.row]
    marks = []
    for r in rows:
        if r.kind != 'section' or not r.label:
            continue
        for kind, words in BLOCK_KEYWORDS:
            if any(w in r.label for w in words):
                marks.append((r.row, kind))
                break

    blocks = []
    for i, (start, kind) in enumerate(marks):
        end = marks[i + 1][0] if i + 1 < len(marks) else 10 ** 6
        body = [r for r in rows if start < r.row < end]
        if body:
            blocks.append((kind, body))
    return blocks


def _rule_block_move(tm, checks):
    """固定资产 / 无形资产 的四块式变动表。

    ① 账面原值  期初 + 本期增加 − 本期减少 = 期末
    ② 累计折旧  同上
    ③ 减值准备  同上
    ④ 账面价值  期末价值 = 原值期末 − 折旧期末 − 减值期末（期初同理）

    这些关系是靠"块标题 + 增加/减少行"定位的，不依赖具体科目名。
    """
    blocks = _find_blocks(tm)
    kinds = {k: body for k, body in blocks}
    if not kinds:
        return

    data_cols = _data_columns(tm)
    if not data_cols:
        return

    def pick(body, *words):
        for r in body:
            if r.label and all(w in r.label for w in words):
                return r
        return None

    ends = {}
    for kind in ('原值', '折旧', '减值'):
        body = kinds.get(kind)
        if not body or len(body) < 3:
            continue
        init_r = body[0]
        end_r = body[-1]
        inc_r = pick(body, '增加')
        dec_r = pick(body, '减少')
        if inc_r is None or dec_r is None:
            continue
        ends[kind] = (init_r, end_r, inc_r, dec_r)

        for c in data_cols:
            checks.append(CheckItem(
                rule_id='R-BLK', level='hard',
                chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
                label=f'{kind}：期末 = 期初 + 本期增加 − 本期减少（{c.header_text or "列"}）',
                target_row=end_r.row, target_col=c.index + 1,
                terms=[Term(+1, init_r.row, c.index + 1),
                       Term(+1, inc_r.row, c.index + 1),
                       Term(-1, dec_r.row, c.index + 1)],
            ))

    # ④ 账面价值 = 原值 − 折旧 − 减值（期末对期末、期初对期初）
    if '价值' in kinds and '原值' in ends:
        body = kinds['价值']
        base_init, base_end = ends['原值'][0], ends['原值'][1]
        init_lab = (base_init.label or '').strip()
        end_lab = (base_end.label or '').strip()

        # 这里的期初期末常写成具体日期（「2026年9月30日账面价值」），
        # 所以先按日期对齐，再退回"期末/期初"字样，最后按常见版式兜底。
        end_v = init_v = None
        for r in body:
            lab = (r.label or '').strip()
            if '期末' in lab or (end_lab and end_lab in lab):
                end_v = end_v or r
            elif '期初' in lab or (init_lab and init_lab in lab):
                init_v = init_v or r
        if end_v is None and body:
            end_v = body[0]
        if init_v is None and len(body) > 1:
            init_v = body[1]

        pairs = []
        for target, idx in ((end_v, 1), (init_v, 0)):
            if target is None or '折旧' not in ends:
                continue
            plus = [ends['原值'][idx]] if '原值' in ends else []
            minus = [ends[k][idx] for k in ('折旧', '减值') if k in ends]
            pairs.append((target, plus, minus))
        for target, plus_rows, minus_rows in pairs:
            for c in data_cols:
                terms = [Term(+1, target.row, c.index + 1)]
                for r in minus_rows:
                    terms.append(Term(-1, r.row, c.index + 1))
                # target 本身不进 terms，右侧是"原值 − 折旧 − 减值"
                terms = [Term(+1, r.row, c.index + 1) for r in plus_rows] + \
                        [Term(-1, r.row, c.index + 1) for r in minus_rows]
                checks.append(CheckItem(
                    rule_id='R-BLK', level='hard',
                    chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
                    label=f'账面价值 = 原值 − 折旧 − 减值（{target.label} / {c.header_text or "列"}）',
                    target_row=target.row, target_col=c.index + 1,
                    terms=terms,
                ))


def _rule_prov_move(tm, checks):
    """坏账/减值准备的变动表：期末 = 期初 + 计提 − 收回或转回 − 转销或核销。

    列名不叫"本期增加/减少"而叫"计提/收回/转销"，所以 R-MOVE 抓不到，
    这是模板里明确写的第三种加减关系。
    """
    cols = _data_columns(tm)
    if len(cols) < 3:
        return

    def col_with(*words):
        for c in cols:
            head = c.header_text or ''
            if all(w in head for w in words):
                return c
        return None

    init_c = col_with('期初')
    end_c = col_with('期末')
    if init_c is None or end_c is None:
        return
    prov_c = col_with('计提')
    back_c = col_with('收回') or col_with('转回')
    write_c = col_with('转销') or col_with('核销')
    other_c = col_with('其他变动')
    if prov_c is None and back_c is None and write_c is None:
        return

    for r in tm.rows:
        if r.kind not in ('detail', 'subtotal') or not r.label:
            continue
        terms = [Term(+1, r.row, init_c.index + 1)]
        if prov_c is not None:
            terms.append(Term(+1, r.row, prov_c.index + 1))
        if back_c is not None:
            terms.append(Term(-1, r.row, back_c.index + 1))
        if write_c is not None:
            terms.append(Term(-1, r.row, write_c.index + 1))
        if other_c is not None:
            v = cell_number(tm.ref.grid, r.row, other_c.index + 1)
            if v:
                terms.append(Term(+1, r.row, other_c.index + 1))
        checks.append(CheckItem(
            rule_id='R-PROV', level='hard',
            chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
            label=f'期末 = 期初 + 计提 − 收回 − 转销（{r.label}）',
            target_row=r.row, target_col=end_c.index + 1,
            terms=terms,
        ))


def _rule_ecl(tm, checks):
    """三阶段预期信用损失变动表（行向）：

        期末余额 = 期初余额 + 期初余额在本期 + 本期计提
                   − 本期转回 − 本期转销 − 本期核销 + 其他变动

    这是模板里出现最多的一种坏账表（81 张、3 份模板），特点是
    「期初余额 / 期初余额在本期 / --转入第二阶段 … / 本期计提 … / 期末余额」竖着排。
    """
    def find(*words, exclude=(), startswith=None):
        for r in tm.rows:
            lab = (r.label or '').strip()
            if not lab:
                continue
            if startswith and not lab.startswith(startswith):
                continue
            if all(w in lab for w in words) and not any(e in lab for e in exclude):
                return r
        return None

    beg = find('期初余额', exclude=('在本期',))
    end = find('期末余额')
    if beg is None or end is None:
        return

    move = find('期初余额在本期')
    prov = find('计提', startswith='本期')
    back = find('转回', startswith='本期')
    write = find('转销', startswith='本期')
    cancel = find('核销', startswith='本期')
    other = find('其他变动')
    if not any((move, prov, back, write, cancel)):
        return

    only_total = len({r.row for r in (beg, end, move, prov, back, write, cancel, other)
                      if r is not None}) <= 4

    for c in _data_columns(tm):
        terms = [Term(+1, beg.row, c.index + 1)]
        if move is not None:
            terms.append(Term(+1, move.row, c.index + 1))
        if prov is not None:
            terms.append(Term(+1, prov.row, c.index + 1))
        if back is not None:
            terms.append(Term(-1, back.row, c.index + 1))
        if write is not None:
            terms.append(Term(-1, write.row, c.index + 1))
        if cancel is not None:
            terms.append(Term(-1, cancel.row, c.index + 1))
        if other is not None:
            terms.append(Term(+1, other.row, c.index + 1))
        checks.append(CheckItem(
            rule_id='R-ECL', level='hard',
            chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
            label=f'期末 = 期初 + 期初在本期 + 计提 − 转回 − 转销 − 核销'
                  f'（{c.header_text or "列"}）',
            target_row=end.row, target_col=c.index + 1,
            terms=terms,
        ))


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
        _rule_block_move(tm, checks)
        _rule_prov_move(tm, checks)
        _rule_ecl(tm, checks)
        _rule_sum_down(tm, checks)
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
