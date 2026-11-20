# -*- coding: utf-8 -*-
"""
report_reader.py — 读三张报表，为"附注 vs 报表"勾稽提供定位能力。

难点
----
1. 一个 Excel 里可能同时有**合并**和**母公司**两套报表 → 靠工作表名区分
2. 科目行的位置每期可能变 → 按**科目名**找行，不写死行号
3. 期间列的位置每期可能变 → 按**表头日期**找列，不写死列号

对外接口
--------
    book = read_report(path)
    book.sheets           # [{name, scope, kind, header_row, periods, items}]
    find_sheet(book, scope, kind)
    find_period_col(sheet, date)
    find_item(sheet, name)
"""

import datetime
import re
from datetime import date
from dataclasses import dataclass, field
from typing import Dict, List, Optional

import openpyxl

RE_DATE = re.compile(r'(\d{4})\s*[年\-/.]\s*(\d{1,2})\s*[月\-/.]\s*(\d{1,2})')
# 「2026年1-9月」这类区间：取期末日
RE_DATE_RANGE = re.compile(
    r'(\d{4})\s*年\s*(\d{1,2})\s*月?\s*[-—~～至到]\s*(\d{1,2})\s*月')
# 「2025年度」「2025年度 收入」：取年末（用 search，不要求整串）
RE_YEAR_ONLY = re.compile(r'(\d{4})\s*年度?(?!\d)')


def _month_end(y, m):
    day = 31 if m in (1, 3, 5, 7, 8, 10, 12) else 30 if m in (4, 6, 9, 11) else 28
    return datetime.date(y, m, day)


def parse_period(value):
    """表头 → 期间日期。支持区间（2026年1-9月 / 2026年1月-9月）、年度（2025年度）、具体日期。

    注意：**区间必须先于日期判断**。像「2026年1-9月」，日期正则会把它吞成
    "2026年1月9日"（2026-01-09），必须先按区间解析成 2026-09-30。
    """
    text = str(value or '').strip()
    if text:
        m = RE_DATE_RANGE.search(text)
        if m:
            y, m2 = int(m.group(1)), int(m.group(3))
            if 1 <= m2 <= 12:
                return _month_end(y, m2)
    d = parse_date(value)
    if d:
        return d
    if not text:
        return None
    m = RE_YEAR_ONLY.search(text)
    if m:
        return datetime.date(int(m.group(1)), 12, 31)
    return None

STATEMENT_KINDS = (
    ('资产负债表', '资产负债表'),
    ('利润表', '利润表'),
    ('损益表', '利润表'),
    ('现金流量表', '现金流量表'),
)
SCOPE_PARENT = ('母公司',)
SCOPE_CONSOL = ('合并',)

# 科目名归一化时要去掉的前后缀
PREFIX_STRIP = ('其中：', '其中:', '减：', '减:', '加：', '加:', '一、', '二、', '三、')
SUFFIX_STRIP = ('（亏损以“－”号填列）', '(亏损以“－”号填列)', '（净亏损以“－”号填列）')


def normalize_name(text):
    """科目名归一化：去空格、去序号、去加减前缀、统一括号。"""
    if text is None:
        return ''
    s = str(text)
    s = s.replace('\u3000', '').replace(' ', '').replace('\n', '')
    s = s.replace('（', '(').replace('）', ')')
    s = s.replace('：', ':')
    changed = True
    while changed:
        changed = False
        for p in PREFIX_STRIP:
            p2 = p.replace('：', ':')
            if s.startswith(p2):
                s = s[len(p2):]
                changed = True
        for suf in SUFFIX_STRIP:
            suf2 = suf.replace('（', '(').replace('）', ')')
            if s.endswith(suf2):
                s = s[:-len(suf2)]
                changed = True
    # 去掉开头的编号，如 "1、" "1." "(1)" "四、" "（二）"
    s = re.sub(r'^[（(]?\d+[)）]?[、.．]?', '', s)
    s = re.sub(r'^[（(]?[一二三四五六七八九十]+[)）]?[、.．]', '', s)
    return s.strip()


def parse_date(value):
    """把单元格解析成日期。

    注意：用 data_only 读 Excel 时，日期格常常拿到的是**序列号**（如 46295），
    不是 datetime，必须按 1899-12-30 起算换算。
    """
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    if value is None:
        return None
    if isinstance(value, (int, float)) and 20000 <= float(value) <= 80000:
        return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(value)))
    m = RE_DATE.search(str(value))
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None


@dataclass
class ReportSheet:
    name: str
    scope: str                     # 合并 / 母公司 / 单体
    kind: str                      # 资产负债表 / 利润表 / 现金流量表
    header_row: int = 0            # 1 基
    periods: Dict[int, object] = field(default_factory=dict)   # 列号(1基) -> date
    items: Dict[str, tuple] = field(default_factory=dict)      # 名称 -> (行号, 科目列)
    raw_names: Dict[tuple, str] = field(default_factory=dict)  # (行号, 科目列) -> 原始行名
    blocks: Dict[int, Dict[int, int]] = field(default_factory=dict)  # 科目列 -> {列号: date}

    def period_columns(self):
        return sorted(self.periods.items())

    def value_col(self, item_col, date):
        """给定科目所在列和日期，返回该科目金额所在列。"""
        block = self.blocks.get(item_col) or {}
        for col, d in block.items():
            if d == date:
                return col
        return None


@dataclass
class ReportBook:
    path: str
    sheets: List[ReportSheet] = field(default_factory=list)

    def find(self, scope, kind):
        for s in self.sheets:
            if s.kind == kind and (scope is None or s.scope == scope):
                return s
        return None


def _scope_of(name):
    if any(k in name for k in SCOPE_PARENT):
        return '母公司'
    if any(k in name for k in SCOPE_CONSOL):
        return '合并'
    return '单体'


def _kind_of(name):
    for key, kind in STATEMENT_KINDS:
        if key in name:
            return kind
    return None


HEADER_WORDS = ('项目', '期末', '年末', '本期', '年初', '期初', '上年', '行次')


def _find_header_row(ws, max_scan=12):
    """找表头行：取"表头证据最强"的一行，而不是第一行。

    报表上方常有标题行/编制单位行，其中可能夹一个日期（如 46022），
    按"第一行有日期"判断会认错。真正的表头行会同时出现
    项目/附注/期末余额/期初余额 这类词，得分更高。
    """
    best, best_score = 0, 0
    for r in range(1, min(max_scan, ws.max_row) + 1):
        dates = words = 0
        for c in range(1, min(ws.max_column, 30) + 1):
            v = ws.cell(r, c).value
            if parse_period(v):
                dates += 1
                continue
            text = str(v or '').strip()
            if len(text) <= 12 and any(w in text for w in HEADER_WORDS):
                words += 1
        if words < 2 and dates < 2:
            continue
        score = dates * 2 + words
        if score > best_score:
            best, best_score = r, score
    return best


RE_NOTE_REF = re.compile(r'^[一二三四五六七八九十百]+[、.]\s*\d')

RE_LATEST_HEAD = ('期末', '本年', '本期', '年末')
RE_PRIOR_HEAD = ('上年', '年初', '期初', '上期')


def _item_columns(ws, sheet, max_col):
    """找出"科目列"：表头下方文字多、不是期间列、也不是「六、1」这类附注号列。"""
    data_start = sheet.header_row + 1
    data_rows = max(0, ws.max_row - data_start + 1)
    item_cols = []
    for c in range(1, max_col + 1):
        if c in sheet.periods:
            continue
        texts = note_refs = 0
        for r in range(data_start, min(ws.max_row, data_start + 300) + 1):
            v = ws.cell(r, c).value
            if not isinstance(v, str):
                continue
            s = v.strip()
            if len(s) < 2:
                continue
            texts += 1
            if RE_NOTE_REF.match(s):
                note_refs += 1
        if texts >= max(5, data_rows * 0.15) and note_refs < texts * 0.5:
            item_cols.append(c)
    return item_cols


def _rebuild_blocks(sheet, item_cols):
    """期间列按「离它左边最近的科目列」归属到分块（资产负债表是左右两栏）。"""
    sheet.blocks = {ic: {} for ic in item_cols}
    for ic in item_cols:
        for col, d in sheet.periods.items():
            if col <= ic:
                continue
            owner = max([x for x in item_cols if x < col], default=None)
            if owner == ic:
                sheet.blocks[ic][col] = d



def _sub_item_rows(ws, item_cols, max_row):
    """找出「其中：」及其续行（靠缩进），这些是明细行，不是独立报表项目。"""
    skip = set()
    for ic in item_cols:
        base_indent = None
        for r in range(1, max_row + 1):
            raw = ws.cell(r, ic).value
            if not isinstance(raw, str):
                continue
            stripped = raw.strip()
            if not stripped:
                continue
            indent = len(raw) - len(raw.lstrip())
            if stripped.startswith("其中"):
                base_indent = indent
                skip.add((r, ic))
                continue
            if base_indent is not None:
                if indent > base_indent:
                    skip.add((r, ic))
                else:
                    base_indent = None
    return skip


def _read_sheet(ws, name):
    kind = _kind_of(name)
    if kind is None:
        return None
    sheet = ReportSheet(name=name, scope=_scope_of(name), kind=kind)
    sheet.header_row = _find_header_row(ws)
    if not sheet.header_row:
        return sheet

    max_col = min(ws.max_column, 40)
    for c in range(1, max_col + 1):
        d = parse_period(ws.cell(sheet.header_row, c).value)
        if d:
            sheet.periods[c] = d

    # 表头只写「期末余额 / 期初余额」而没有具体日期时，用标题行里的日期
    # （如「2025年度」「2025年12月31日」）推出各列期间
    if not sheet.periods:
        doc_date = None
        for r in range(1, sheet.header_row):
            for c in range(1, max_col + 1):
                d = parse_period(ws.cell(r, c).value)
                if d and (doc_date is None or d > doc_date):
                    doc_date = d
        if doc_date:
            prior = date(doc_date.year - 1, 12, 31)
            for c in range(1, max_col + 1):
                text = str(ws.cell(sheet.header_row, c).value or '').strip()
                if not text:
                    continue
                if any(k in text for k in RE_PRIOR_HEAD):
                    sheet.periods[c] = prior
                elif any(k in text for k in RE_LATEST_HEAD):
                    sheet.periods[c] = doc_date

    item_cols = _item_columns(ws, sheet, max_col)
    _rebuild_blocks(sheet, item_cols)

    data_start = sheet.header_row + 1
    # 不再丢「其中：」行 —— 它们是报表的一部分，分类表会把它们标成"子项"。
    # 以前在这里整行丢掉，导致子项（其中：应收利息、其中：应付利息…）
    # 根本没进报表模型，和原始版对不上。
    skip = set()
    for ic in item_cols:
        for r in range(data_start, ws.max_row + 1):
            if (r, ic) in skip:
                continue          # 「其中：」子项，不是独立报表项目
            raw = ws.cell(r, ic).value
            key = normalize_name(raw)
            if key and key not in sheet.items:
                sheet.items[key] = (r, ic)
                if isinstance(raw, str):
                    sheet.raw_names[(r, ic)] = raw.strip()
    sheet._item_cols = item_cols
    return sheet


def _borrow_periods(wb, book):
    """表头只写「期末余额 / 上年年末余额」的报表，向有日期的兄弟表借用期间。"""
    donor = next((s for s in book.sheets if s.period_columns()), None)
    if donor is None:
        return
    ordered = [d for _, d in donor.period_columns()]
    if len(ordered) < 1:
        return

    for sheet in book.sheets:
        if sheet.period_columns() or not sheet.header_row:
            continue
        ws = wb[sheet.name]
        max_col = min(ws.max_column, 40)
        for c in range(1, max_col + 1):
            text = str(ws.cell(sheet.header_row, c).value or '').strip()
            if not text:
                continue
            # 先判"上年/年初"，再判"期末/本年"——"上年年末余额"里也含"年末"，
            # 顺序反了会被错当成最近一期。
            if any(k in text for k in RE_PRIOR_HEAD) and len(ordered) > 1:
                sheet.periods[c] = ordered[1]
            elif any(k in text for k in RE_LATEST_HEAD):
                sheet.periods[c] = ordered[0]
        if not sheet.periods:
            continue
        _rebuild_blocks(sheet, getattr(sheet, '_item_cols', []))


def read_report(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    book = ReportBook(path=path)
    for name in wb.sheetnames:
        ws = wb[name]
        sheet = _read_sheet(ws, name)
        if sheet is not None:
            book.sheets.append(sheet)
    # 表头只写「期末余额/上年年末余额」的报表，借用兄弟表的期间
    _borrow_periods(wb, book)
    return book


def find_item(sheet, name):
    """按科目名找位置，返回 (行号, 科目列) 或 None。先全等，再包含匹配。"""
    if sheet is None:
        return None
    key = normalize_name(name)
    if not key:
        return None
    if key in sheet.items:
        return sheet.items[key]
    best = None
    for item, pos in sheet.items.items():
        if not item:
            continue
        if len(key) >= 3 and (key in item or item in key):
            if best is None or abs(len(item) - len(key)) < abs(len(best[0]) - len(key)):
                best = (item, pos)
    return best[1] if best else None
