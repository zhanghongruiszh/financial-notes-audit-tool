# -*- coding: utf-8 -*-
"""
disclosure_checks.py — 报表项目在 Word 附注中「有没有披露」。

用途
----
报表上某个科目有金额，但附注主要项目里找不到对应的小节，这可能是：
  · 真的漏披露（要补）
  · 本来就不需要披露（比如某些报表科目在附注里并入别的项目）
所以这里**只做汇总、不下结论**，输出一张清单供人工判断。

判定口径
--------
· 只考虑**有金额**的报表科目（各期全为 0 或「-」的，谈不上披露）
· 跳过小计/合计/总计行（它们由明细构成，本身不是披露对象）
· 科目名匹配**只做全等与去括号后的全等**，不做包含匹配——
  宁可多列一条让人看一眼，也不要把没披露的说成已披露
"""

import re
from dataclasses import dataclass, field

import openpyxl

import report_numbering as rn
import report_reader as rr

SKIP_WORDS = ('合计', '总计', '小计')

# 报表上的"算出来的"项目，本身不是披露对象（附注不会给它单独一节）
SKIP_ITEMS = {
    '营业利润', '利润总额', '净利润', '综合收益总额',
    '持续经营净利润', '终止经营净利润',
    '基本每股收益', '稀释每股收益',
    '流动资产合计', '非流动资产合计', '资产总计',
    '流动负债合计', '非流动负债合计', '负债合计',
    '所有者权益合计', '负债和所有者权益总计',
}

# 现金流量表的每个项目在附注里通常不单独披露（补充资料由 TB 核对覆盖），
# 为避免整表刷屏，默认不纳入"未披露"判断
SKIP_SHEET_KINDS = ('现金流量表',)

NIL = {'', '-', '--', '—', '——', '–', '/'}


@dataclass
class DisclosureItem:
    sheet: str
    row: int
    item_name: str
    amounts: dict = field(default_factory=dict)   # {date: value}
    item_col: int = 0        # 所在分栏（资产负债表左右两栏分开排序）
    matched: str = ''
    status: str = 'missing'
    note: str = ''


def _to_number(raw):
    if isinstance(raw, bool):
        return None
    if isinstance(raw, (int, float)):
        return float(raw)
    text = str(raw or '').strip().replace(',', '')
    if text in NIL:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return None


def _note_name_set(model, scope):
    """该范围下附注子项目的归一化名字集合。

    除了子项目全名，还拆出「和/、/及」分隔的各部分——
    附注写「营业收入和营业成本」，报表上是「营业收入」「营业成本」两行，
    不拆就会把已披露的误判成未披露。
    """
    chapter = rn._pick_chapter(model, scope)
    if chapter is None:
        return set()
    names = set()
    for sub in chapter.sub_items:
        raw = sub.name
        names.add(rr.normalize_name(raw))
        stripped = rr.normalize_name(rn._strip_brackets(raw))
        if stripped:
            names.add(stripped)
        for part in re.split(r'[和、及与]', raw):
            part = part.strip()
            if len(part) >= 2:
                names.add(rr.normalize_name(part))
    return names


def _sub_item_rows(ws, rows, item_col):
    """找出属于「其中：」块的行（子项）。

    报表里子项有两种写法，都要排除——它们不是报表项目，只是上一行的拆分，
    附注里在上级项目下披露即可，不需要单独小节：
        · 本行以「其中：」开头
        · 本行缩进比上面那个「其中：」更深（续行，如「利息收入」「应付股利」）
    """
    sub = set()
    base_indent = None
    for r in sorted(rows):
        raw = str(ws.cell(r, item_col).value or '')
        stripped = raw.strip()
        if not stripped:
            continue
        indent = len(raw) - len(raw.lstrip())
        if stripped.startswith('其中'):
            base_indent = indent
            sub.add(r)
            continue
        if base_indent is not None:
            if indent > base_indent:
                sub.add(r)        # 「其中：」块里的续行
            else:
                base_indent = None
    return sub


def build_disclosure_checks(report_path, book, model):
    """返回有金额、但附注里找不到对应小节的报表科目清单。"""
    wb = openpyxl.load_workbook(report_path, data_only=True)

    out = []
    for sheet in book.sheets:
        if any(k in sheet.name for k in SKIP_SHEET_KINDS):
            continue
        names = _note_name_set(model, sheet.scope)
        if not names or not sheet.period_columns():
            continue
        if sheet.name not in wb.sheetnames:
            continue
        ws = wb[sheet.name]

        # 该表的「其中：」子项行，不参与"未披露"判断
        col_rows = {}
        for key, (row, item_col) in sheet.items.items():
            col_rows.setdefault(item_col, []).append(row)
        skip_rows = set()
        # 只有资产负债表才把「其中/缩进」当子项跳过；
        # 利润表里「税金及附加」这类是正式报表项目，必须参与未披露判断
        if '资产负债表' in sheet.name:
            for item_col, rows in col_rows.items():
                skip_rows |= _sub_item_rows(ws, rows, item_col)

        seen = set()
        for key, (row, item_col) in sheet.items.items():
            if not key or key in seen:
                continue
            if row in skip_rows:
                continue          # 「其中：」子项，附注里在上级项目下披露即可
            if any(w in key for w in SKIP_WORDS):
                continue
            # 合计性质/子项性质的项目不需要单独披露
            if rn.is_no_note_item(key):
                continue
            # 与别的项目共用同一附注号的（如"实收资本净额"），不算漏披露
            if key in {rr.normalize_name(k) for k in rn.REPORT_ITEM_ALIAS}:
                continue
            key_nb = rr.normalize_name(rn._strip_brackets(key))
            if key_nb in SKIP_ITEMS or key in SKIP_ITEMS:
                continue

            amounts = {}
            for value_col, date in (sheet.blocks.get(item_col) or {}).items():
                v = _to_number(ws.cell(row, value_col).value)
                if v is not None:
                    amounts[date] = v
            if not amounts or all(abs(v) < 0.005 for v in amounts.values()):
                continue                      # 各期都是 0，谈不上披露
            seen.add(key)

            if key in names or (key_nb and key_nb in names):
                continue                      # 已披露

            out.append(DisclosureItem(
                sheet=sheet.name, row=row, item_name=key, amounts=amounts,
                status='missing',
                note='附注未找到同名小节（可能是漏披露，也可能无需披露）',
            ))

    out.sort(key=lambda x: (x.sheet, x.row))
    return out


def summarize(items):
    return {'missing': len(items)}
