# -*- coding: utf-8 -*-
"""
tb_checks.py — 附注「现金流量表补充资料」与 TB 现金流量表的逐项核对。

为什么单独做
------------
对外三张报表只列现金流量表的合计数，附注补充资料里的间接法调节项、
现金及现金等价物构成等，报表上没有完整对应项目，必须回 TB 底稿取数。
"""

import datetime
from dataclasses import dataclass
from typing import Optional

import openpyxl

import rules
import tb_reader

# 附注用词 与 TB 用词的差异（只补常见几组，其余靠"包含匹配"）
ALIAS = {
    '信用减值损失': ['信用减值准备', '信用减值损失'],
    '资产减值准备': ['资产减值准备', '资产减值损失'],
    '投资性房地产折旧': ['投资性房地产折旧'],
}

# 说明行/汇总行不参与逐项核对（它们由 ①②③ 分节标题带过）
SKIP_LABELS = ('将净利润调节为经营活动现金流量', '不涉及现金收支的重大投资和筹资活动',
               '现金及现金等价物净变动情况')


@dataclass
class TbCheck:
    chapter: str
    sub_item: str
    table_idx: int
    label: str
    period: Optional[datetime.date]
    note_value: Optional[float]
    note_ref: str = ''
    note_row: int = 0              # 附注表内行号（用于写到底稿右侧）
    note_col: int = 0
    tb_sheet: str = ''
    tb_rows: Optional[int] = None
    tb_cols: Optional[int] = None
    tb_value: Optional[float] = None
    diff: Optional[float] = None
    status: str = 'unmatched'      # pass / fail / unmatched
    note: str = ''


def _is_cf_subitem(name):
    """只有「现金流量表补充资料」这张表与 TB 的 CF-Indirect 一一对应。

    货币资金下的「可随时用于支付的银行存款」等，虽然名字像现流，
    但它们属于货币资金附注，TB 的 CF 表里没有对应行，不能拿来勾。
    """
    n = name or ''
    return '现金流量表' in n and '货币资金' not in n


SYNONYMS = {
    '资产减值损失': ['资产减值准备', '计提资产减值准备'],
    '资产减值准备': ['资产减值损失'],
}


def _candidates(label):
    """给一个附注行文字，返回可尝试匹配的名字列表。"""
    text = label.strip()
    for pre in ('加：', '加:', '减：', '减:'):
        if text.startswith(pre):
            text = text[len(pre):]
            break
    out = [label] + SYNONYMS.get(label.strip(), []) + SYNONYMS.get(text, [])
    for key, alts in ALIAS.items():
        if key in label:
            out.extend(alts)
    return out


def _col_date(col, cf):
    """表头是「本期发生额/上期发生额」这种没有具体日期的，按语义对到 TB 的期间。

    TB 现金流量表一般有两年（本年/上年）：本期对最近一期、上期对上一期。
    """
    if col.period:
        return col.period
    periods = sorted(cf.periods.keys(), reverse=True)
    if not periods:
        return None
    text = (col.header_text or '') + (col.period_key or '')
    if any(w in text for w in ('本期', '本年', '期末')):
        idx = 0
    elif any(w in text for w in ('上期', '上年', '期初', '年初')):
        idx = 1
    else:
        return None
    return periods[idx] if idx < len(periods) else None


def build_tb_checks(model, tables, tb_sheets):
    cf = tb_reader.find_cf(tb_sheets, prefer='indirect')
    if cf is None:
        return []

    out = []
    for tm in tables:
        if not _is_cf_subitem(tm.sub_item):
            continue
        # 只认"净利润调节表"：必须同时有净利润、经营活动现金流量净额两行。
        # 「现金及现金等价物构成」那张小表（可随时用于支付的存款等）不在 TB 里，
        # 不参与 TB 勾稽。
        labels = [(r.label or '') for r in tm.rows]
        if not (any('净利润' in l for l in labels)
                and any('经营活动产生的现金流量净额' in l for l in labels)):
            continue
        for row in tm.rows:
            if row.kind != 'detail' or not row.label:
                continue
            if any(k in row.label for k in SKIP_LABELS):
                continue
            for col in tm.columns:
                if not col.is_data:
                    continue
                date = _col_date(col, cf)
                if date is None:
                    continue
                note_val = rules.cell_number(tm.ref.grid, row.row, col.index + 1)
                if note_val is None:
                    continue

                chk = TbCheck(
                    chapter=tm.chapter, sub_item=tm.sub_item, table_idx=tm.ref.idx,
                    label=row.label, period=date, note_value=note_val,
                    note_ref=f'表#{tm.ref.idx} R{row.row} C{col.index + 1}',
                    note_row=row.row, note_col=col.index + 1,
                )

                tb_row = None
                for cand in _candidates(row.label):
                    tb_row = tb_reader.find_item(cf, cand)
                    if tb_row:
                        break
                if not tb_row:
                    chk.status = 'unmatched'
                    chk.note = 'TB 现金流量表里没找到同名项目'
                    out.append(chk)
                    continue

                tb_col = cf.value_col(date)
                chk.tb_sheet = cf.name
                chk.tb_rows = tb_row
                if tb_col is None:
                    chk.status = 'unmatched'
                    chk.note = f'TB 里没有 {col.period} 这一期'
                    out.append(chk)
                    continue
                chk.tb_cols = tb_col
                out.append(chk)
    return out


def _to_number(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if v is None:
        return 0.0
    s = str(v).strip().replace(',', '')
    if s in ('', '-', '--', '—', '——', '–', '/'):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return None


def evaluate(checks, tb_xlsx_path):
    """用 TB 工作簿里的实际值填充并判定。"""
    wb = openpyxl.load_workbook(tb_xlsx_path, data_only=True)
    for chk in checks:
        if not (chk.tb_sheet and chk.tb_rows and chk.tb_cols):
            continue
        if chk.tb_sheet not in wb.sheetnames:
            continue
        raw = wb[chk.tb_sheet].cell(chk.tb_rows, chk.tb_cols).value
        chk.tb_value = _to_number(raw)
        if chk.tb_value is None:
            chk.status = 'unmatched'
            chk.note = f'TB 取数不是数字：{raw!r}'
            continue
        chk.diff = chk.note_value - chk.tb_value
        chk.status = 'pass' if abs(chk.diff) <= rules.TOLERANCE else 'fail'
    return checks
