# -*- coding: utf-8 -*-
"""
tb_reader.py — 读 TB 试算平衡表里的现金流量表（CF）工作表。

为什么需要
----------
Word 附注里的「现金流量表补充资料」在对外报表上没有完整对应项目，
必须去 TB 底稿的 CF 表取审定数来核对。

TB 的版式
---------
以本项目为例，`CF-Indirect` 表：

    第6行: 项目名称 | 2023年 | SAD |    | 2023年 |    | 2024年 | ...
    第7行:          | （B.A）|  +  | -  | （A.A）|    | （B.A）| ...

每个年度占 **5 列**：`（B.A）  +  -  （A.A）  空`。
对外披露用的是 **（A.A）审定数** 那一列。

所以本模块做两件事：
  1. 认出"年度块"，并定位每个年度的（A.A）列
  2. 把列 A 的科目名归一化建索引，供按名字找行

.xls 无法用 openpyxl 读，先用 Excel COM 转成 .xlsx 并缓存。
"""

import datetime
import os
import re

import openpyxl

import report_reader as rr

WORK_DIR = '_work/tb'

RE_AA = re.compile(r'[（(]\s*A\.?\s*A\.?\s*[)）]', re.IGNORECASE)
RE_BA = re.compile(r'[（(]\s*B\.?\s*A\.?\s*[)）]', re.IGNORECASE)


def ensure_xlsx(path, work_dir=None):
    """把 .xls 转成 .xlsx（缓存），返回 .xlsx 路径。"""
    path = os.path.abspath(path)
    if path.lower().endswith('.xlsx'):
        return path
    work_dir = os.path.abspath(work_dir or WORK_DIR)
    os.makedirs(work_dir, exist_ok=True)
    dst = os.path.join(work_dir, os.path.splitext(os.path.basename(path))[0] + '.xlsx')
    if os.path.exists(dst) and os.path.getmtime(dst) >= os.path.getmtime(path):
        return dst

    import win32com.client as win32
    import com_env
    import docx_ir

    src_use, _ = com_env.ascii_staging(path, '.xls', work_dir)
    early = com_env.staging_dir(work_dir)
    if early:
        import hashlib
        tag = hashlib.md5(os.path.abspath(path).encode('utf-8')).hexdigest()[:10]
        out_stage = os.path.join(early, f'tb_{tag}.xlsx')
    else:
        out_stage = dst

    def _convert():
        with com_env.sanitized():
            xl = win32.DispatchEx('Excel.Application')
        try:
            xl.Visible = False
            xl.DisplayAlerts = False
            wb = xl.Workbooks.Open(os.path.abspath(src_use), ReadOnly=True)
            wb.SaveAs(os.path.abspath(out_stage), FileFormat=51)   # xlOpenXMLWorkbook
            wb.Close(False)
        finally:
            try:
                xl.Quit()
            except Exception:
                pass
        return out_stage

    last_error = None
    for attempt in (1, 2):
        try:
            docx_ir._run_com_with_timeout(_convert, 120, 'Excel 转换 TB')
            last_error = None
            break
        except Exception as exc:      # noqa: BLE001
            last_error = exc
            if attempt == 1:
                # Office 首次启动会初始化缓存，个别环境第一次必失败，再试一次
                import time
                time.sleep(2)
    if last_error is not None:
        raise last_error
    if os.path.abspath(out_stage) != os.path.abspath(dst):
        import shutil
        shutil.copy2(out_stage, dst)
    return dst


_CN_DIGIT = {'1': '一', '2': '二', '3': '三', '4': '四', '5': '五',
             '6': '六', '7': '七', '8': '八', '9': '九', '0': '零'}


def normalize_label(text):
    """科目名归一化：去空格、去括号注释、去"加：/减："前缀、个位数字转中文。"""
    if text is None:
        return ''
    s = str(text).replace('\u3000', '').replace(' ', '').strip()
    s = s.replace('（', '(').replace('）', ')')
    s = re.sub(r'[\(（][^)）]*[)）]', '', s)      # 去掉括号注释
    s = re.sub(r'^(加:|减:)', '', s)
    s = re.sub(r'^\d+[.、]', '', s)
    # 「1年内到期的可转换公司债券」vs TB「一年内到期…」：个位数字转中文
    s = re.sub(r'(?<!\d)([0-9])(?=年)', lambda m: _CN_DIGIT.get(m.group(1), m.group(1)), s)
    return s.strip()


class CfSheet:
    """TB 里的一张现金流量表。"""

    def __init__(self, name):
        self.name = name
        self.header_row = 0
        self.periods = {}      # date -> 审定数(A.A)列号
        self.items = {}        # 归一化名称 -> 行号
        self.label_col = 1

    def value_col(self, date):
        return self.periods.get(date)


def _find_header_row(ws, max_scan=15):
    for r in range(1, min(max_scan, ws.max_row) + 1):
        years = 0
        for c in range(1, min(ws.max_column, 40) + 1):
            if rr.parse_period(ws.cell(r, c).value):
                years += 1
        if years >= 2:
            return r
    return 0


def _read_cf_sheet(ws, name):
    sheet = CfSheet(name)
    sheet.header_row = _find_header_row(ws)
    if not sheet.header_row:
        return None

    hr = sheet.header_row
    max_col = min(ws.max_column, 60)

    # 第 1 步：找出每个"年度块"的起始列
    blocks = []
    for c in range(1, max_col + 1):
        d = rr.parse_period(ws.cell(hr, c).value)
        if d:
            blocks.append((c, d))

    # 第 2 步：在每个块里定位（A.A）列
    for start, d in blocks:
        aa_col = None
        for c in range(start, min(start + 6, max_col + 1)):
            text = str(ws.cell(hr + 1, c).value or '')
            if RE_AA.search(text):
                aa_col = c
                break
        if aa_col is not None:
            sheet.periods[d] = aa_col

    if not sheet.periods:
        return None

    # 第 3 步：科目名（列 A）建索引
    for r in range(hr + 2, ws.max_row + 1):
        key = normalize_label(ws.cell(r, 1).value)
        if key and key not in sheet.items:
            sheet.items[key] = r
    return sheet


def read_tb(path, work_dir=None):
    """返回 {工作表名: CfSheet}，只包含现金流量表相关的表。"""
    xlsx = ensure_xlsx(path, work_dir)
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    out = {}
    for name in wb.sheetnames:
        if 'CF' not in name.upper():
            continue
        sheet = _read_cf_sheet(wb[name], name)
        if sheet is not None:
            out[name] = sheet
    return out


def find_cf(tb_sheets, prefer='indirect'):
    """挑一张 CF 表：默认优先间接法（补充资料用）。"""
    if not tb_sheets:
        return None
    for name, s in tb_sheets.items():
        if prefer == 'indirect' and 'indirect' in name.lower():
            return s
    for name, s in tb_sheets.items():
        if prefer == 'direct' and 'direct' in name.lower():
            return s
    return next(iter(tb_sheets.values()))


def find_item(sheet, label, min_len=4):
    """按归一化名称找 TB 行，先全等再包含匹配。"""
    key = normalize_label(label)
    if not key:
        return None
    if key in sheet.items:
        return sheet.items[key]
    best = None
    for item, row in sheet.items.items():
        if not item:
            continue
        if len(key) >= min_len and (key in item or item in key):
            if best is None or abs(len(item) - len(key)) < abs(len(best[0]) - len(key)):
                best = (item, row)
    return best[1] if best else None
