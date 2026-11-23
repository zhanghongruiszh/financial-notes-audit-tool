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
    alts: List[List['Term']] = None                   # 备选口径（任一成立即算平）
    alt_den: List['Term'] = None                      # 比例类备选分母
    soft_fail: bool = False                           # 备选也不成立时，标存疑而不是差异
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

    if item.level == 'unclear':
        return CheckResult(item, expected, None, None, 'unclear',
                           '结构存疑：认不准哪几行属于这个合计，请人工确认')

    actual = cell_number(grid, item.target_row, item.target_col)
    if actual is None:
        # 空白视为 0：空模板/无金额的表也要有公式勾稽（0 = 0 即平）
        actual = 0.0
    diff = actual - expected

    # 口径不唯一时（重分类方向、前五名分母选哪张表），先试备选口径
    if item.alt_den and item.is_ratio and abs(diff) * 100 > RATIO_TOLERANCE_PP:
        try:
            alt_d = sum(t.coef * (_val(t) or 0.0) for t in item.alt_den)
            if abs(alt_d) > 1e-9:
                alt_exp = total / alt_d
                if abs(actual - alt_exp) * 100 <= RATIO_TOLERANCE_PP:
                    return CheckResult(item, alt_exp, actual, actual - alt_exp, 'pass')
        except Exception:
            pass

    if item.alts and abs(diff) > TOLERANCE:
        for alt in item.alts:
            try:
                alt_exp = sum(t.coef * (_val(t) or 0.0) for t in alt)
            except Exception:
                continue
            if item.den:
                continue
            if abs(actual - alt_exp) <= TOLERANCE:
                return CheckResult(item, alt_exp, actual, actual - alt_exp, 'pass')
    if item.is_ratio:
        # 比例按"百分点"比较：0.01 个百分点以内算平
        if abs(diff) * 100 <= RATIO_TOLERANCE_PP:
            status = 'pass'
        elif item.soft_fail:
            status = 'unclear'      # 分母口径不唯一，不硬判差异
        else:
            status = 'fail'
        if item.level == 'hard':
            item.level = 'soft'
    elif abs(diff) <= TOLERANCE:
        status = 'pass'
    elif item.soft_fail:
        status = 'unclear'
    else:
        status = 'fail'
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

def _candidate_blocks(tm, total_row, ref_col):
    """给合计行生成几个"候选行组"，用数值验证挑一个。

    光靠语义会把「其中」行算错，光靠数值会被重述行带偏；
    两者结合：语义出候选，数值只做**验证**，都不成立就标结构存疑。
    """
    rows = {r.row: r for r in tm.rows}
    # 向上收集同级行（遇到分节/合计就停）
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
            block.append(r)
            break
        block.append(r)
        r -= 1
    block.sort()

    def value(row_no):
        return cell_number(tm.ref.grid, row_no, ref_col) or 0.0

    base = []          # 候选一：排除其中/减项，含小计
    with_break = []    # 候选二：把"父行为空"的其中行也算进来
    subtotal_first = []  # 候选三：只取小计/合计行
    for row_no in block:
        info = rows[row_no]
        label = (info.label or '').strip()
        is_break = info.kind == 'group' or label.startswith(('其中', '减：', '减:'))
        if info.kind in ('total', 'subtotal'):
            base.append(row_no)
            subtotal_first.append(row_no)
            with_break.append(row_no)
            continue
        if is_break:
            continue
        base.append(row_no)
        with_break.append(row_no)
        subtotal_first.append(row_no)

    # 候选二：父行为空的「其中」行补进来
    last_parent = None
    for row_no in block:
        info = rows[row_no]
        label = (info.label or '').strip()
        is_break = info.kind == 'group' or label.startswith(('其中', '减：', '减:'))
        if not is_break:
            last_parent = row_no
            continue
        if last_parent is not None and abs(value(last_parent)) < 0.005:
            with_break.append(row_no)

    # 候选四：把紧跟「其中：」后面的续行当作该组的子项排除掉
    no_group_children = []
    after_group = False
    for row_no in block:
        info = rows[row_no]
        label = (info.label or '').strip()
        is_break = info.kind == 'group' or label.startswith(('其中', '减：', '减:'))
        if is_break:
            after_group = True
            continue
        if after_group and info.kind == 'detail':
            continue        # 视为上一组的子项，不计入合计
        after_group = False
        no_group_children.append(row_no)

    cands = []
    for cand in (base, no_group_children, with_break, subtotal_first):
        cand = sorted(set(cand))
        if cand and cand not in cands:
            cands.append(cand)
    return cands




def _rule_tax_split(tm, checks):
    """其他综合收益这类"税前 − 所得税 = 税后净额"的横向关系。

    全按表头关键字识别（税前/所得税/税后），不认表名、不认固定列号。
    """
    pre = tax = post = None
    for c in tm.columns:
        head = (c.header_text or '').replace(' ', '')
        if not head:
            continue
        if pre is None and ('税前' in head or '税前金额' in head):
            pre = c
        elif tax is None and '所得税' in head:
            tax = c
        elif post is None and ('税后' in head):
            post = c
    if not (pre and post):
        return

    for r in tm.rows:
        if r.kind not in ('detail', 'subtotal', 'total') or not r.label:
            continue
        if not cell_number(tm.ref.grid, r.row, post.index + 1):
            continue
        terms = [Term(+1, r.row, pre.index + 1)]
        if tax is not None:
            terms.append(Term(-1, r.row, tax.index + 1))
        checks.append(CheckItem(
            rule_id='R-TAX', level='hard',
            chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
            label='税后净额 = 税前金额 − 所得税',
            target_row=r.row, target_col=post.index + 1,
            terms=terms,
        ))



def _rule_lease(tm, checks):
    """租赁负债：净额 = 租赁付款额 − 未确认的融资费用 − 重分类至一年内到期。

    只看行名，不依赖表类型（会计准则口径固定）。
    """
    rows = [r for r in tm.rows
            if r.kind in ('detail', 'subtotal', 'total') and r.label]

    def find(*words):
        for r in rows:
            if all(w in (r.label or '') for w in words):
                return r
        return None

    pay = find('租赁付款额')
    fee = find('未确认的融资费用')
    # 三个要素必须齐：付款额 + 未确认融资费用（一年内到期可缺）
    if pay is None or fee is None:
        return
    due = find('一年内到期')
    net = find('净额') or find('租赁负债')

    target = net if net is not None else None
    if target is None:
        return
    for c in _data_columns(tm):
        terms = [Term(+1, pay.row, c.index + 1)]
        if fee is not None:
            terms.append(Term(-1, fee.row, c.index + 1))
        if due is not None:
            terms.append(Term(-1, due.row, c.index + 1))
        checks.append(CheckItem(
            rule_id='R-LEASE', level='hard',
            chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
            label=f'租赁负债净额 = 付款额 − 未确认融资费用 − 一年内到期（{c.header_text or "列"}）',
            target_row=target.row, target_col=c.index + 1,
            terms=terms,
        ))


def _rule_asset_blocks(tm, checks):
    """固定资产/无形资产/使用权资产这类"分段式"表的纵向加减：

        账面净值合计 = 账面原值合计 − 累计折旧(摊销)合计
        账面价值合计 = 账面净值合计 − 减值准备合计

    这几行都是 subtotal（不是 total），普通 V-SUM 抓不到。
    """
    rows = [r for r in tm.rows if r.kind in ('subtotal', 'total') and r.label]

    def find(*words):
        for r in rows:
            if all(w in r.label for w in words):
                return r
        return None

    gross = find('原值') or find('原价')
    dep = find('累计折旧') or find('累计摊销')
    net = find('净值')
    imp = find('减值准备')
    value = find('账面价值')

    # 会计准则口径（不依赖表类型，只要有这几行就必须成立）：
    #   账面净值 = 原值 − 累计折旧(摊销)
    #   账面价值 = 账面净值 − 减值准备（减值没有也按 0 处理）
    if not (gross and dep):
        return
    if value is None and net is None:
        return

    def _is_movement(c):
        head = (c.header_text or '')
        return any(w in head for w in ('增加', '减少', '重分类', '变动'))

    # (12) 每段小计 = 其下明细之和（原值/折旧/减值/净值/价值 各自分段）
    blocks = [r for r in tm.rows
              if r.kind in ('total', 'subtotal') and r.label
              and any(w in r.label for w in ('原值', '原价', '折旧', '摊销',
                                             '净值', '减值', '账面价值'))]
    detail_rows = [r for r in tm.rows
                   if r.kind in ('detail', 'group') and r.label]
    for i, blk in enumerate(blocks):
        end_row = blocks[i + 1].row if i + 1 < len(blocks) else 10 ** 6
        kids = [r for r in detail_rows if blk.row < r.row < end_row]
        if len(kids) < 2:
            continue
        for c in _data_columns(tm):
            if not cell_number(tm.ref.grid, blk.row, c.index + 1):
                continue
            checks.append(CheckItem(
                rule_id='R-ABLK', level='hard',
                chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
                label=f'{blk.label} = 其下 {len(kids)} 项之和（{c.header_text or "列"}）',
                target_row=blk.row, target_col=c.index + 1,
                terms=[Term(row_sign(r.label), r.row, c.index + 1) for r in kids],
            ))

    for c in _data_columns(tm):
        if value is not None and _is_movement(c):
            # 净值/账面价值这两行只有期初期末，没有发生数，不核发生列
            continue
        if value is None:
            # 只到净值这一层：原值 − 折旧 = 净值
            g = cell_number(tm.ref.grid, gross.row, c.index + 1)
            d = cell_number(tm.ref.grid, dep.row, c.index + 1)
            if (g or d) and net is not None:
                checks.append(CheckItem(
                    rule_id='R-ABLK', level='hard',
                    chapter=tm.chapter, sub_item=tm.sub_item,
                    table_idx=tm.ref.idx,
                    label=f'账面净值 = 原值 − 累计折旧(摊销)（{c.header_text or "列"}）',
                    target_row=net.row, target_col=c.index + 1,
                    terms=[Term(+1, gross.row, c.index + 1),
                           Term(-1, dep.row, c.index + 1)],
                ))
            continue
        if net is not None:
            # 原值、折旧都有数，净值这一列却是空的 → 净值漏填，必须报
            g = cell_number(tm.ref.grid, gross.row, c.index + 1)
            d = cell_number(tm.ref.grid, dep.row, c.index + 1)
            if g and d and not cell_number(tm.ref.grid, net.row, c.index + 1):
                checks.append(CheckItem(
                    rule_id='R-ABLK', level='hard',
                    chapter=tm.chapter, sub_item=tm.sub_item,
                    table_idx=tm.ref.idx,
                    label=f'账面净值漏填（{c.header_text or "列"}）',
                    target_row=net.row, target_col=c.index + 1,
                    terms=[Term(+1, gross.row, c.index + 1),
                           Term(-1, dep.row, c.index + 1)],
                ))
            # 净值/价值那几行只有期初、期末有数，增加/减少列是空的 → 跳过
            if net.kind and cell_number(tm.ref.grid, net.row, c.index + 1):
                checks.append(CheckItem(
                    rule_id='R-ABLK', level='hard',
                    chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
                    label=f'账面净值 = 原值 − 累计折旧（{c.header_text or "列"}）',
                    target_row=net.row, target_col=c.index + 1,
                    terms=[Term(+1, gross.row, c.index + 1),
                           Term(-1, dep.row, c.index + 1)],
                ))
            # 注意：净值行**存在但为空**是"漏填"，必须报差异，不能拿别的口径替它算
            # 只有整张表根本没有"净值"行时，才退回 原值 − 折旧 − 减值
            has_net = True          # 结构上存在净值行
            base = net
        else:
            has_net = False
            base = gross
        if not cell_number(tm.ref.grid, value.row, c.index + 1):
            continue
        # 账面价值这一列有数、但净值行是空的 → 这是"净额漏填"，必须报差异；
        # 只有账面价值本身为空时才跳过（那一列本来就没有勾稽关系）
        if not cell_number(tm.ref.grid, value.row, c.index + 1):
            continue
        # 净值行为空时，"漏填"已经单独报过，这里不再重复报价值不符
        if net is not None and not cell_number(tm.ref.grid, net.row, c.index + 1):
            continue
        terms = [Term(+1, base.row, c.index + 1)]
        if not has_net:
            terms.append(Term(-1, dep.row, c.index + 1))   # 整表没有净值行才多减一次
        if imp is not None:
            terms.append(Term(-1, imp.row, c.index + 1))
        checks.append(CheckItem(
            rule_id='R-ABLK', level='hard',
            chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
            label=f'账面价值 = 原值 − 折旧 − 减值（{c.header_text or "列"}）',
            target_row=value.row, target_col=c.index + 1,
            terms=terms,
        ))


def _rule_v_sum(tm, checks):
    """合计行 = 上方明细之和（语义候选 + 数值验证，逐列判定）。

    行组按"语义"生成几个候选，再用每一列的数值验证挑一个：
      · 某列能对上 → 用它（该列按"确定"处理）
      · 这一列一个候选都对不上，但别的列能对上 → 结构存疑（黄色）
      · 所有列都对不上 → 真的平不了（红色差异）
    """
    grid = tm.ref.grid

    # 带「税前 / 所得税 / 税后」三列的表（其他综合收益这类），
    # 主要关系是横向的"税后 = 税前 − 所得税"；纵向合并不做，
    # 否则分段小计会被误当加总行。
    heads = [(c.header_text or '') for c in tm.columns if c.is_data]
    if any('税前' in h for h in heads) and any('税后' in h for h in heads):
        return

    # 有些表（如按公司横向分列的投资明细）没有 total 行，只有「资产合计」这类小计；
    # 分段式资产表的小计（账面原值合计/累计折旧合计…）不能当加总行，要排除
    SEGMENT_WORDS = ('原值', '原价', '折旧', '摊销', '净值', '减值', '账面价值')
    totals = [r for r in tm.rows if r.kind == 'total']
    if not totals:
        totals = [r for r in tm.rows
                  if r.kind == 'subtotal' and '合计' in (r.label or '')
                  and not any(w in (r.label or '') for w in SEGMENT_WORDS)]
    if not totals:
        return
    cols = _data_columns(tm, include_ratio=False)
    rows = {r.row: r for r in tm.rows}

    for tot in totals:
        ref_col = _pick_ref_col(tm, tot.row, cols)
        cands = _candidate_blocks(tm, tot.row, ref_col)
        # 补一个候选：只取"上一个合计行之后"的明细
        # （负债合计 = 流动负债 + 非流动负债，不能把上面的资产段也算进来）
        above = [r for r in tm.rows
                 if r.row < tot.row and r.label]
        tail = []
        for r in reversed(above):
            if r.kind in ('total', 'subtotal'):
                break
            if r.kind in ('detail', 'group'):
                tail.append(r)
        if tail:
            tail = [r.row for r in reversed(tail)]
            if tail not in cands:
                cands = [tail] + list(cands)

            # 「其中：xxx」组：组本身和紧跟其后的若干明细都是子项，
            # 不能进合计。组后接几行不确定，逐个尝试，由数值验证挑。
            group_pos = [i for i, r in enumerate(tm.rows)
                         if r.row in tail and r.kind == 'group']
            for gi in group_pos:
                for k in range(0, 5):
                    drop = {tm.rows[gi].row}
                    j = gi + 1
                    cnt = 0
                    while j < len(tm.rows) and cnt < k:
                        if tm.rows[j].row in tail:
                            drop.add(tm.rows[j].row)
                            cnt += 1
                        j += 1
                    variant = [x for x in tail if x not in drop]
                    if variant and variant not in cands:
                        cands.append(variant)

        legacy = _contribution_rows(tm, tot.row, ref_col)
        if legacy:
            cands = [legacy] + [c for c in cands if c != legacy]
        if not cands:
            continue

        def acc_of(cand, col):
            return sum(row_sign(rows[r].label)
                       * (cell_number(grid, r, col) or 0.0) for r in cand)

        # 逐列找能对上的候选
        picks = {}
        for c in cols:
            target = cell_number(grid, tot.row, c.index + 1)
            if not target and c.role is None:
                # 合计行这一列是空的，且这列本来就不是金额列（债券面值/日期/期限），
                # 不算加总关系，不生成勾稽
                continue
            found = None
            flip = False
            for cand in cands:
                if target is None:
                    continue
                acc = acc_of(cand, c.index + 1)
                if abs(acc - target) <= TOLERANCE:
                    found, flip = cand, False
                    break
                # 有些列整体按"减项"呈现，合计是负的
                if abs(acc + target) <= TOLERANCE:
                    found, flip = cand, True
                    break
            picks[c.index] = (found, flip)

        cands = sorted(cands, key=len)
        any_found = any(v[0] for v in picks.values())
        for c in cols:
            if c.index not in picks:
                continue
            found, flip = picks[c.index]
            if found is not None:
                level = 'hard'
                picked = found
            elif not cell_number(grid, tot.row, c.index + 1):
                # 合计行这一列是空的：按"一切合计都要做加减"的口径，算差异
                level = 'hard'
                picked = cands[0]
                flip = False
            elif any_found or len(cols) > 1:
                # 已经能给出期望值，就按差异报，让用户看到具体数，
                # 不再用"结构存疑"兜底（用户明确不喜欢成片黄色）
                level = 'hard'
                picked = cands[0]
                flip = False
            else:
                level = 'hard'         # 单列表还平不了 → 真的差异
                picked = cands[0]
                flip = False

            coef_dir = -1 if flip else 1
            terms = [Term(coef_dir * row_sign(rows[r].label), r, c.index + 1)
                     for r in picked]
            checks.append(CheckItem(
                rule_id='V-SUM', level=level,
                chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
                label=f'合计行 = 上方 {len(picked)} 行之和'
                      f'（{c.header_text or "列" + str(c.index + 1)}）'
                      + ('　[结构存疑]' if level == 'unclear' else ''),
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
            # 父行这一列写的是「——」/「—」这类占位符时，说明该列不是加总关系；
            # 但**空白**不能跳——空着的父行照样要核对子项
            try:
                cell = tm.ref.grid[parent_row - 1][c.index]
                raw_text = str(cell.get('t') or '').strip()
            except Exception:
                raw_text = ''
            if raw_text in ('——', '—', '--', '-', '–', '/'):
                continue
            # 有的表把子项写到"其中"块下面一格（合并格延续），
            # 子项和为 0 而父行有数时，把紧跟块尾的下一行也试进来
            kids = list(children)
            csum0 = sum(row_sign(rows[r].label)
                        * (cell_number(tm.ref.grid, r, c.index + 1) or 0.0)
                        for r in kids)
            pv0 = cell_number(tm.ref.grid, parent_row, c.index + 1) or 0.0
            if not csum0 and pv0 and kids:
                nxt = max(kids) + 1
                if nxt < parent_row and nxt not in kids:
                    kids.append(nxt)
                    try:
                        rows.setdefault(nxt, next(r for r in tm.rows if r.row == nxt))
                    except StopIteration:
                        kids.pop()
            children = kids
            child_sum = sum(row_sign(rows[r].label)
                            * (cell_number(tm.ref.grid, r, c.index + 1) or 0.0)
                            for r in children)
            # 子项在别处（块外延续行）时，期望值按 0 算，直接报差异，
            # 不再用黄色兜底
            level = 'hard'
            terms = [Term(row_sign(rows[r].label), r, c.index + 1) for r in children]
            checks.append(CheckItem(
                rule_id='V-GRP', level=level,
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
            # 增加类：全部计入；减少类：全部计入（含"摊销/折旧"这种减少项）；
            # 中间的"重分类/其他"不动——它方向不定，算进去反而错。
            extras = [c for c in cols
                      if dec.index < c.index < end_col.index
                      and c.role not in ('增加', '减少', '余额', '价值', '准备', '比例')]
            plus_inc = [c for c in cols if c.role == '增加']
            minus_dec = [c for c in cols if c.role in ('减少', '折旧')]
            # 净值/账面价值块里的行是倒算出来的，不服从"期初+增加−减少"
            value_block_rows = set()
            cur_is_value = False
            for r in tm.rows:
                label = (r.label or '')
                if r.kind in ('section', 'subtotal') and any(
                        k in label for k in ('净值', '账面价值')):
                    cur_is_value = True
                    continue
                if cur_is_value and r.kind in ('detail', 'subtotal', 'total'):
                    value_block_rows.add(r.row)

            for r in tm.rows:
                if r.kind not in ('detail', 'total', 'subtotal') or not r.label:
                    continue
                if any(k in (r.label or '') for k in ('账面价值', '净值', '净额')):
                    continue
                if r.row in value_block_rows:
                    continue
                terms = [Term(+1, r.row, beg_col.index + 1)]
                terms += [Term(+1, r.row, c.index + 1) for c in plus_inc]
                terms += [Term(-1, r.row, c.index + 1) for c in minus_dec]
                has_extra = any((cell_number(tm.ref.grid, r.row, c.index + 1) or 0)
                                for c in extras)
                checks.append(CheckItem(
                    rule_id='R-MOVE', level='hard',
                    chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
                    label=f'期末 = 期初 + 本期增加 − 本期减少（{r.label}）',
                    target_row=r.row, target_col=end_col.index + 1,
                    terms=terms,
                    alts=[[Term(+1, r.row, beg_col.index + 1)]
                          + [Term(+1, r.row, c.index + 1) for c in plus_inc]
                          + [Term(-1, r.row, c.index + 1) for c in minus_dec]
                          + [Term(+1, r.row, c.index + 1) for c in extras]]
                    if extras else
                    [[Term(+1, r.row, beg_col.index + 1)]
                     + [Term(+1, r.row, c.index + 1) for c in plus_inc]
                     + [Term(+1, r.row, c.index + 1) for c in minus_dec]],
                    soft_fail=bool(has_extra),
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
            # 用户口径：占比分母取"该子项目**第一张**表的合计数"，
            # 不是合计最大的那张（分类表合计可能比账龄表大，会算小占比）
            if best is None or other.ref.idx < best[0].ref.idx:
                best = (other, tot, col, val)
        return best[:3] if best else None

    total_rows = [r for r in tm.rows if r.kind == 'total']
    tot = total_rows[-1] if total_rows else None
    targets = [r for r in tm.rows if r.kind in ('detail', 'subtotal') and r.label]
    company_rows = sum(1 for r in tm.rows
                       if r.label and any(w in r.label for w in ('有限公司', '股份', '集团')))
    is_top5 = (getattr(tm, 'table_type', '') == 'TOP5'
               or company_rows >= 3)

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
            # 先用表内数据反推分母（金额 ÷ 占比），再去候选表里找最接近的那个：
            # 详情表常把"账面余额合计"和"账面价值合计"并排列示，
            # 占比用的是前者，光按期间选列会取错。
            implied = None
            for r0 in tm.rows:
                if r0.kind not in ('detail', 'subtotal') or not r0.label:
                    continue
                amt0 = cell_number(tm.ref.grid, r0.row, base.index + 1)
                pct0 = cell_number(tm.ref.grid, r0.row, rc.index + 1)
                if amt0 and pct0:
                    implied = amt0 / pct0
                    break
            found = None
            if implied:
                best = None
                for other in (siblings or []):
                    if other is tm or other.ref.idx >= tm.ref.idx:
                        continue
                    tot0 = next((r for r in other.rows
                                 if r.kind in ('total', 'subtotal')), None)
                    if tot0 is None:
                        continue
                    for c0 in other.columns:
                        if not c0.is_data:
                            continue
                        v0 = cell_number(other.ref.grid, tot0.row, c0.index + 1)
                        if not v0:
                            continue
                        d = abs(v0 - implied) / abs(implied)
                        if best is None or d < best[0]:
                            best = (d, other, tot0, c0)
                if best and best[0] <= 0.01:
                    found = (best[1], best[2], best[3])
            if found is None:
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
                    alt_den=[Term(+1, tot.row, base.index + 1)] if tot else None,
                    soft_fail=True,
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
    if getattr(tm, 'table_type', '') not in ('PROVISION', 'STAGED_PROVISION',):
        return
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
    if getattr(tm, 'table_type', '') not in ('STAGED_PROVISION',):
        return
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
        # 三阶段表的变动行本身就是**带正负号的金额**（计提可能是负数、
        # 核销也常写成负数），所以一律相加，不能再套一次符号。
        terms = [Term(+1, beg.row, c.index + 1)]
        for r in (move, prov, back, write, cancel, other):
            if r is not None:
                terms.append(Term(+1, r.row, c.index + 1))
        checks.append(CheckItem(
            rule_id='R-ECL', level='hard',
            chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
            label=f'期末 = 期初 + 期初在本期 + 计提 − 转回 − 转销 − 核销'
                  f'（{c.header_text or "列"}）',
            target_row=end.row, target_col=c.index + 1,
            terms=terms,
        ))


# 整段不做表内勾稽的子项目（按用户口径）
SKIP_SUB_EXACT = {'租赁'}
SKIP_SUB_WORDS = ('不重要合营', '不重要联营', '不重要子', '汇总信息')


def _skip_table(tm):
    """这些表不做表内勾稽：

      · 「不重要合营企业和联营企业的汇总信息」这类汇总表 —— 只与报表勾
      · 「租赁」整个段落 —— 定性披露为主
      · 只有一个数据行的表 —— 除了和报表勾没有别的公式
    """
    name = str(getattr(tm, 'sub_item', '') or '')
    if name in SKIP_SUB_EXACT:
        return True
    if any(w in name for w in SKIP_SUB_WORDS):
        return True
    details = [r for r in tm.rows
               if r.kind in ('detail', 'subtotal', 'total') and r.label]
    has_total = any(r.kind in ('total', 'subtotal') for r in tm.rows)
    # 真正的"单格表"：只有一个数据格（一行一列），才没有表内勾稽可言
    cells = 0
    for r in details:
        for c in tm.columns:
            if c.is_data and cell_number(tm.ref.grid, r.row, c.index + 1):
                cells += 1
    if cells <= 1 and not has_total:
        return True
    return False




def _rule_restricted_money(tables, checks):
    """受限资产表里若有「货币资金」这一行，
    该行金额应与「货币资金」项目下受限子表的合计数一致。
    """
    def find_row(tm, *words):
        for r in tm.rows:
            if r.kind in ('detail', 'subtotal', 'total') and r.label:
                if all(w in r.label for w in words):
                    return r
        return None

    money_tables = [t for t in tables if '货币资金' in str(t.sub_item)]

    for tm in tables:
        if '受限' not in str(tm.sub_item):
            continue
        row = find_row(tm, '货币资金')
        if row is None:
            continue
        # 货币资金项目下、含"受限/监管/冻结"的那张表
        src = None
        for mt in money_tables:
            hits = [r for r in mt.rows if r.label and any(
                w in r.label for w in ('受限', '监管', '冻结', '涉诉'))]
            if hits:
                src = mt
                break
        if src is None:
            continue
        total = find_row(src, '合计')
        if total is None:
            continue
        for c in _data_columns(tm):
            checks.append(CheckItem(
                rule_id='R-RESTRICT', level='hard',
                chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
                label='受限货币资金 = 货币资金受限子表合计',
                target_row=row.row, target_col=c.index + 1,
                terms=[Term(+1, total.row, c.index + 1, table_idx=src.ref.idx)],
            ))



# 拆成两张表时，列名的增减含义（靠关键字判断，不看列角色）
_SPLIT_INC = ('本年发行', '发行金额', '计提利息', '溢折价摊销', '摊销', '本年增加')
_SPLIT_DEC = ('本年偿还', '偿还金额', '本年减少', '收回金额')


def _rule_split_move(tables, checks):
    """拆成两张表的变动表：期末（第二张） = 期初（第一张） + 增加 − 减少。

    典型：应付债券 / 其他流动负债 —— 第一张有期初，第二张有本年发行、
    计提利息、摊销、偿还、期末。
    """
    def col_by(tm, *words, role=None):
        for c in tm.columns:
            if not c.is_data:
                continue
            head = (c.header_text or '').replace(' ', '')
            if any(w in head for w in words):
                return c
        return None

    prev = None
    for tm in tables:
        if prev is not None and prev.sub_item == tm.sub_item:
            a, b = prev, tm
            end = col_by(b, '期末余额') or col_by(b, '期末')
            beg = col_by(a, '期初余额') or col_by(a, '期初')
            if end is not None and beg is not None:
                incs = [c for c in b.columns if c.is_data and c is not end
                        and any(w in (c.header_text or '') for w in _SPLIT_INC)]
                decs = [c for c in b.columns if c.is_data and c is not end
                        and any(w in (c.header_text or '') for w in _SPLIT_DEC)]
                a_rows = {(r.label or '').strip(): r for r in a.rows
                          if r.kind in ('detail', 'subtotal', 'total') and r.label}

                # 整表验证（看总额，不看逐行）：所有行加总若不平，
                # 说明这两张表根本不是"期初→期末"的关系，整对跳过。
                # 逐行的偏差留待下面的检查去发现。
                tot_end = tot_beg = tot_inc = tot_dec = 0.0
                paired = 0
                for rb in b.rows:
                    if rb.kind not in ('detail', 'subtotal', 'total') or not rb.label:
                        continue
                    ra = a_rows.get((rb.label or '').strip())
                    if ra is None:
                        continue
                    paired += 1
                    tot_end += cell_number(b.ref.grid, rb.row, end.index + 1) or 0.0
                    tot_beg += cell_number(a.ref.grid, ra.row, beg.index + 1) or 0.0
                    for c in incs:
                        tot_inc += cell_number(b.ref.grid, rb.row, c.index + 1) or 0.0
                    for c in decs:
                        tot_dec += cell_number(b.ref.grid, rb.row, c.index + 1) or 0.0
                if paired < 2 or abs(tot_end - (tot_beg + tot_inc - tot_dec)) > 0.01:
                    prev = tm
                    continue

                for rb in b.rows:
                    if rb.kind not in ('detail', 'subtotal', 'total') or not rb.label:
                        continue
                    ra = a_rows.get((rb.label or '').strip())
                    if ra is None:
                        continue
                    if not cell_number(b.ref.grid, rb.row, end.index + 1):
                        continue
                    terms = [Term(+1, ra.row, beg.index + 1, table_idx=a.ref.idx)]
                    terms += [Term(+1, rb.row, c.index + 1) for c in incs]
                    terms += [Term(-1, rb.row, c.index + 1) for c in decs]
                    checks.append(CheckItem(
                        rule_id='R-SPLIT', level='hard',
                        chapter=tm.chapter, sub_item=tm.sub_item,
                        table_idx=b.ref.idx,
                        label=f'期末 = 期初 + 增加 − 减少（{rb.label}·续表）',
                        target_row=rb.row, target_col=end.index + 1,
                        terms=terms,
                    ))
        prev = tm



# 递延所得税税率（企业所得税法定税率）
DEFERRED_TAX_RATE = 0.25


def _rule_deferred_tax(tm, checks):
    """递延所得税：所得税金额 = 可抵扣/应纳税暂时性差异 × 25%。

    表头成对出现：「递延所得税资产/负债」列 与「可抵扣/应纳税暂时性差异」列，
    按期间前缀（期末/期初）配对。
    """
    cols = [c for c in tm.columns if c.is_data]
    pairs = []
    for i, c in enumerate(cols):
        head = (c.header_text or '').replace(' ', '')
        if '暂时性差异' not in head:
            continue
        # 同一期间前缀的"递延所得税"列
        for d in cols:
            h2 = (d.header_text or '').replace(' ', '')
            if '递延所得税' not in h2:
                continue
            pre = head.split('可抵扣')[0].split('应纳税')[0]
            if h2.startswith(pre):
                pairs.append((d, c))
                break
    if not pairs:
        return

    for r in tm.rows:
        if r.kind not in ('detail', 'subtotal', 'total') or not r.label:
            continue
        for tax_col, diff_col in pairs:
            diff = cell_number(tm.ref.grid, r.row, diff_col.index + 1)
            if not diff:
                continue
            checks.append(CheckItem(
                rule_id='R-DTAX', level='hard',
                chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
                label=f'递延所得税 = 暂时性差异 × 25%（{r.label}）',
                target_row=r.row, target_col=tax_col.index + 1,
                terms=[Term(+DEFERRED_TAX_RATE, r.row, diff_col.index + 1)],
                is_ratio=False,
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
        if _skip_table(tm):
            continue
        _rule_sum_down(tm, checks)
        _rule_asset_blocks(tm, checks)
        _rule_deferred_tax(tm, checks)
        _rule_lease(tm, checks)
        _rule_tax_split(tm, checks)
        _rule_pct(tm, checks, siblings)
    try:
        _rule_restricted_money(tables, checks)
        _rule_split_move(tables, checks)
    except Exception:
        pass
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
