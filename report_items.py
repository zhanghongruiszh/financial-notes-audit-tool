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
    """判断一行是 报表项目 / 子项 / 合计项，以及子项的归属。"""
    if terms.is_total(name):
        return 'total', None
    parent = terms.parent_of(name)
    if parent:
        return 'sub', parent
    if any(n and _norm(name) == _norm(n) for n in terms.OCI_LINES):
        return 'sub', '其他综合收益'
    text = str(name or '').strip()
    if text.endswith((':', '：')) and len(text) <= 12:
        return 'section', None      # 「流动资产:」这类分段小标题
    if text.startswith(('其中', '减：', '减:', '加：', '加:')):
        return 'sub', prev_line
    return 'line', None


def ordered_items(book):
    """把三张表规范化成一条有序的处理清单。

    返回 [ReportItem]：
      · 先按 SHEET_ORDER 排工作表（不依赖文件里的物理顺序）
      · 每张表先左段（项目列靠左）再右段
      · 段内按行号从上到下
    """
    out = []
    sheets = sorted(book.sheets, key=lambda s: sheet_rank(s.name))
    order = 0
    for sheet in sheets:
        # 分栏：按项目列号分组，左段在前
        blocks = {}
        for name, (row, item_col) in sheet.items.items():
            blocks.setdefault(item_col, []).append((row, name))
        for item_col in sorted(blocks):
            rows = sorted(blocks[item_col])
            prev_line = None
            for row, name in rows:
                kind, parent = classify(name, prev_line)
                out.append(ReportItem(sheet.name, item_col, row, item_col,
                                      name, kind, parent, order))
                order += 1
                if kind == 'line':
                    prev_line = name
    return out


def line_items(book):
    """只要报表项目（参与和 Word 标题匹配的那些）。"""
    return [it for it in ordered_items(book) if it.kind == 'line']


def sub_items(book):
    """只要子项（它们的归属 parent 就是要去表里找行的那个项目）。"""
    return [it for it in ordered_items(book) if it.kind == 'sub']


def totals(book):
    return [it for it in ordered_items(book) if it.kind == 'total']
