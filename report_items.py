# -*- coding: utf-8 -*-
"""
report_items.py — 报表项目的"规范化数据处理"层。

用户明确要求的顺序与分类：

  1. 工作表顺序：资产负债表 → 利润表 → 现金流量表 → 所有者权益表
     （没有的表跳过；顺序不依赖文件里 sheet 的物理顺序）
  2. 每张表分**左右两段**，先左段（资产/收入）后右段（负债权益/支出）
  3. 每一行按性质分成两类：
       · 报表项目 : 真正的独立科目 —— 与 Word **主要附注的标题**匹配
       · 子项     : 某个项目的组成部分（其中：xxx、实收资本净额、OCI 各子项、
                    合计数下面的明细）—— 与对应标题**下方表格的行**匹配
     合计项（营业总收入、资产总计…）既不是报表项目也不是子项，
     它只是加总结果，不参与匹配。

匹配思路（本模块只负责"排序 + 分类 + 找到归属"，不做数值比对）：
    报表项目 → 找同名的 Word 小节标题（或按报表附注号定位）
    子项     → 在该小节的表里找同名行
"""

import re

import report_item_terms as terms

# 工作表优先级：靠前的先处理
SHEET_ORDER = ('资产负债表', '利润表', '现金流量表', '所有者权益')


def _norm(text):
    s = re.sub(r'[（(][^）)]*[）)]', '', str(text or ''))
    return re.sub(r'\s+', '', s)


def sheet_rank(name):
    for i, key in enumerate(SHEET_ORDER):
        if str(name).startswith(key):
            return i
    return len(SHEET_ORDER)


class ReportItem:
    """报表里的一行（规范化后）。"""

    __slots__ = ('sheet', 'block', 'row', 'col', 'name', 'kind', 'parent',
                 'order')

    def __init__(self, sheet, block, row, col, name, kind, parent, order):
        self.sheet = sheet      # 工作表名
        self.block = block      # 分栏（左段/右段），用项目列号表示
        self.row = row
        self.col = col          # 项目所在列
        self.name = name
        self.kind = kind        # 'line' 报表项目 / 'sub' 子项 / 'total' 合计项
        self.parent = parent    # 子项所属的报表项目
        self.order = order      # 处理顺序（sheet → 左/右段 → 行）

    def __repr__(self):
        return f'<{self.sheet} {self.kind} {self.name} P={self.parent}>'


def classify(name, prev_line=None):
    """判断一行是 报表项目 / 子项 / 合计项，以及子项的归属。

    优先查**人工分类表**（规则库/report_item_classes.json），
    查不到再走启发式。
    """
    cls = terms.class_of(name)
    if cls == '合计项':
        return 'total', None
    if cls == '报表项目':
        return 'line', None
    if cls == '子项':
        text = str(name or '').strip()
        if text.startswith(('其中', '减：', '减:', '加：', '加:')):
            return 'sub', prev_line
        return 'sub', terms.parent_of(name) or prev_line
    if cls == '分类标题':
        return 'title', None

    if terms.is_total(name):
        return 'total', None
    parent = terms.parent_of(name)
    if parent:
        return 'sub', parent
    if any(n and _norm(name) == _norm(n) for n in terms.OCI_LINES):
        return 'sub', '其他综合收益'
    text = str(name or '').strip()
    if text.endswith((':', '：')) and len(text) <= 12:
        return '分类标题', None      # 「流动资产：」这类分类标题
    if text.startswith(('其中', '减：', '减:', '加：', '加:')):
        return 'sub', prev_line
    return 'line', None


def _raw_indent(wb, sheet_name, row, col):
    """读原始单元格文字的缩进（Word/Excel 里靠前导空格表示层级）。"""
    if wb is None or sheet_name not in getattr(wb, 'sheetnames', []):
        return 0
    try:
        raw = wb[sheet_name].cell(row, col).value
    except Exception:
        return 0
    if not isinstance(raw, str):
        return 0
    return len(raw) - len(raw.lstrip())


def ordered_items(book, wb=None):
    """把三张表规范化成一条有序的处理清单。

    顺序：工作表（资产负债表→利润表→现流表→权益表）→ 左段→右段 → 行号
    分类代码：total / line / sub / title / template，
    显示名与人工分类表一致（合计项 / 报表项目 / 子项 / 分类标题）
      · 合计项只看术语表（营业总收入、营业总成本、资产总计…）
      · 子项优先看术语表里的 SUB_ITEMS（实收资本净额、利息费用、OCI 各子项…），
        其次看行名前缀「其中/加：/减：」
      · 其余都是报表项目
    缩进在客户文件里并不可靠（营业成本缩进 3、税金及附加缩进 9，
    但它们是同级），所以**不用缩进推断层级**，一律以术语表为准。
    """
    out = []
    sheets = sorted(book.sheets, key=lambda s: sheet_rank(s.name))
    order = 0
    for sheet in sheets:
        blocks = {}
        for name, (row, item_col) in sheet.items.items():
            blocks.setdefault(item_col, []).append((row, name))
        _sheet_items_start = len(out)
        for item_col in sorted(blocks):
            rows = sorted(blocks[item_col])
            prev_line = None
            for row, name in rows:
                text = str(name or '').strip()
                # 人工分类表把 △/▲/☆ 这些也当作报表项目（只是本期无金额），
                # 所以只有当分类表查不到时才退回"模板行"
                if (text.startswith(('△', '▲', '☆', '#'))
                        and terms.class_of(name) is None):
                    out.append(ReportItem(sheet.name, item_col, row, item_col,
                                          name, 'template', None, order))
                    order += 1
                    continue
                if text.endswith((':', '：')) and len(text) <= 12:
                    out.append(ReportItem(sheet.name, item_col, row, item_col,
                                          name, 'title', None, order))
                    order += 1
                    continue
                kind, parent = classify(name, prev_line)
                out.append(ReportItem(sheet.name, item_col, row, item_col,
                                      name, kind, parent, order))
                order += 1
                if kind == 'line':
                    prev_line = name
        _fix_sub_parents(out, _sheet_items_start, sheet.name)
    return out


def _fix_sub_parents(items, start, sheet_name):
    """子项的归属：资产负债表/利润表往上找，现金流量表往下找。

    同方向找不到时再用反方向兜底。
    """
    seg = items[start:]
    down = str(sheet_name).startswith('现金流量表')

    def _lines():
        return [i for i in seg if i.kind == 'line']

    lines = _lines()
    for idx, it in enumerate(seg):
        if it.kind != 'sub':
            continue
        parent = None
        # 现金流量表里"投资/筹资活动产生的现金流量净额"被归为合计项，
        # 但它们正是子项的归属锚点，所以两种类型都算候选
        anchors = ('line', 'total') if down else ('line',)
        if down:
            for nxt in seg[idx + 1:]:
                if nxt.kind in anchors and nxt.block == it.block:
                    parent = nxt.name
                    break
        else:
            for prv in reversed(seg[:idx]):
                if prv.kind in anchors and prv.block == it.block:
                    parent = prv.name
                    break
        if parent is None:      # 反方向兜底（两种类型都看）
            for other in (seg[idx + 1:] if not down else reversed(seg[:idx])):
                if other.kind in ('line', 'total') and other.block == it.block:
                    parent = other.name
                    break
        if parent:
            it.parent = parent


def line_items(book):
    """只要报表项目（参与和 Word 标题匹配的那些）。"""
    return [it for it in ordered_items(book) if it.kind == 'line']


def sub_items(book):
    """只要子项（它们的归属 parent 就是要去表里找行的那个项目）。"""
    return [it for it in ordered_items(book) if it.kind == 'sub']


def totals(book):
    return [it for it in ordered_items(book) if it.kind == 'total']
