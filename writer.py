# -*- coding: utf-8 -*-
"""
writer.py — 把勾稽结果写成 Excel 成果表（3 个工作表）。

工作表结构（按需求收拢，避免一堆 sheet 影响可读性）
----------------------------------------------------
1. **勾稽底稿**（主体）
   与 Word 附注一致的顺序与版式逐个表铺开：
     · 勾稽点**紧贴目标行下方**插入一行【勾稽校验】，写活公式
     · 表格**最右侧**接报表勾稽列：报表数（粘贴）+ 差额（活公式 = 附注 − 报表）
       现金流量表补充资料那张表右侧再接 TB 核对列
2. **汇总表**
   勾稽总览（各核对类型条数与结果）+ 内部勾稽的差异清单
3. **完整性**
   未披露汇总、附注号勾稽、附注内跨表——都是清单性质，放一起

为什么报表数用"粘贴值 + 差额公式"
----------------------------------
不把整份报表复制进工作簿，是为了不让 sheet 数量膨胀。报表数原样贴在右侧独立列，
差额用公式算，点开就能看到"附注哪一格 − 报表哪一格"，可追溯性不丢。
"""

import datetime
from collections import Counter

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import rules

# ---------------------------------------------------------------- 样式

F_TITLE = Font(name='微软雅黑', size=10, bold=True, color='FFFFFF')
FILL_TITLE = PatternFill('solid', start_color='1F497D', end_color='1F497D')

F_TABLE_TITLE = Font(name='微软雅黑', size=9, bold=True, color='1F497D')
FILL_TABLE_TITLE = PatternFill('solid', start_color='DCE6F1', end_color='DCE6F1')

F_HEADER = Font(name='微软雅黑', size=9, bold=True, color='333333')
FILL_HEADER = PatternFill('solid', start_color='F2F4F8', end_color='F2F4F8')

F_BODY = Font(name='微软雅黑', size=9, color='000000')
F_CHECK = Font(name='微软雅黑', size=9, color='1F497D')
F_REPORT = Font(name='微软雅黑', size=9, color='7030A0')

FILL_PASS = PatternFill('solid', start_color='E2EFDA', end_color='E2EFDA')
FILL_FAIL_TARGET = PatternFill('solid', start_color='FFC7CE', end_color='FFC7CE')
FILL_FAIL_ROW = PatternFill('solid', start_color='FDE2E2', end_color='FDE2E2')
FILL_UNKNOWN = PatternFill('solid', start_color='FFF2CC', end_color='FFF2CC')
FILL_CHECK_ROW = PatternFill('solid', start_color='F7F9FC', end_color='F7F9FC')
FILL_REPORT_HDR = PatternFill('solid', start_color='E8DFF5', end_color='E8DFF5')
FILL_TB_HDR = PatternFill('solid', start_color='DDEBF7', end_color='DDEBF7')

THIN = Side(style='thin', color='D9D9D9')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FMT_AMOUNT = '#,##0.00;[Red](#,##0.00);"-"'
FMT_RATIO = '0.00%;[Red]-0.00%;"-"'
FMT_DATE = 'yyyy-mm-dd'

REPORT_COL_START = 11          # K 列起：报表勾稽
TB_COL_START_OFFSET = 2        # 报表块之后接 TB 块


# ---------------------------------------------------------------- 基础写入

def _write_value(ws, row, col, value):
    if value is None:
        return
    cell = ws.cell(row, col, value)
    cell.font = F_BODY
    cell.border = BORDER
    if isinstance(value, datetime.date):
        cell.number_format = FMT_DATE
        cell.alignment = Alignment(horizontal='center')
    elif isinstance(value, float):
        cell.number_format = FMT_AMOUNT
        cell.alignment = Alignment(horizontal='right')
    else:
        cell.alignment = Alignment(horizontal='left', vertical='center')


def _column_is_ratio(tm, col_index):
    for c in tm.columns:
        if c.index == col_index and c.kind == 'ratio':
            return True
    return False


def _period_slots(cross, tb_checks, limit=3):
    """右侧列的期间顺序（最近的一期在最左）。"""
    periods = []
    for cc in (cross or []):
        if cc.period and cc.period not in periods:
            periods.append(cc.period)
    for t in (tb_checks or []):
        if t.period and t.period not in periods:
            periods.append(t.period)
    periods.sort(reverse=True)
    return periods[:limit]


def _col_map(periods):
    """每期只留一列「差额」公式，原始数字直接在公式里引用「原始数据」工作表。"""
    report = {}
    tb = {}
    for i, p in enumerate(periods):
        report[p] = REPORT_COL_START + i
        base = REPORT_COL_START + len(periods) + TB_COL_START_OFFSET
        tb[p] = base + i
    return report, tb


def _write_right_headers(ws, row, periods, report_map, tb_map, has_tb):
    for p in periods:
        c = ws.cell(row, report_map[p], f'报表差额 {p}')
        c.font = F_REPORT
        c.fill = FILL_REPORT_HDR
        c.alignment = Alignment(horizontal='center')
        c.border = BORDER
        if has_tb:
            c = ws.cell(row, tb_map[p], f'TB差额 {p}')
            c.font = F_REPORT
            c.fill = FILL_TB_HDR
            c.alignment = Alignment(horizontal='center')
            c.border = BORDER


# ---------------------------------------------------------------- 底稿

def _write_table(ws, tm, checks_by_row, results_by_key, row_ptr, stats,
                 global_rows, report_by_row, tb_by_row, report_map, tb_map, raw=None,
                 check_rows=None, anomalies=None):
    grid = tm.ref.grid

    # 科目名放在第一格（与 Word 里的标题一致、显眼），技术信息放右边小字
    title_cell = ws.cell(row_ptr, 1, tm.sub_item)
    title_cell.font = Font(name='微软雅黑', size=11, bold=True, color='1F497D')
    title_cell.fill = FILL_TABLE_TITLE
    info_cell = ws.cell(row_ptr, max(3, tm.ref.cols),
                        f'表#{tm.ref.idx}　{tm.ref.rows}行×{tm.ref.cols}列　'
                        f'{tm.table_type_name or "未分类"}')
    info_cell.font = Font(name='微软雅黑', size=8, color='808080')
    info_cell.fill = FILL_TABLE_TITLE
    info_cell.alignment = Alignment(horizontal='right')
    for c in range(2, max(3, tm.ref.cols) + 1):
        ws.cell(row_ptr, c).fill = FILL_TABLE_TITLE
    row_ptr += 1

    hard = [w for w in tm.warnings if not w.startswith('（期间待确认）')]
    soft = [w for w in tm.warnings if w.startswith('（期间待确认）')]
    if hard:
        cell = ws.cell(row_ptr, 1, '⚠ ' + '；'.join(hard[:3]))
        cell.font = Font(name='微软雅黑', size=9, color='B25E00')
        cell.fill = FILL_UNKNOWN
        stats['unclear_tables'] += 1
        row_ptr += 1
    if soft:
        cell = ws.cell(row_ptr, 1, '· ' + '；'.join(
            w.replace('（期间待确认）', '') for w in soft[:4]))
        cell.font = Font(name='微软雅黑', size=8, color='808080')
        row_ptr += 1

    src_to_dst = {}
    note_row_map = {}
    for r_idx, row in enumerate(grid, 1):
        for c_idx, cellinfo in enumerate(row, 1):
            if cellinfo['m'] != ' ':
                continue
            text = cellinfo['t']
            if not text:
                continue
            import docx_ir
            if anomalies is not None and docx_ir.is_trailing_minus(text):
                # 「9,900.00-」这种尾随负号：数值照算，但要提示把 - 删掉
                anomalies.append({
                    'sub_item': tm.sub_item, 'table': tm.ref.idx,
                    'row': r_idx, 'col': c_idx, 'text': str(text).strip(),
                })
            value = docx_ir.parse_cell_value(text)
            _write_value(ws, row_ptr, c_idx, value)
            if _column_is_ratio(tm, c_idx - 1) and isinstance(value, float):
                ws.cell(row_ptr, c_idx).number_format = FMT_RATIO
        kind = next((r.kind for r in tm.rows if r.row == r_idx), None)
        if kind == 'header':
            for c_idx in range(1, tm.ref.cols + 1):
                cell = ws.cell(row_ptr, c_idx)
                cell.font = F_HEADER
                cell.fill = FILL_HEADER
        src_to_dst[r_idx] = row_ptr
        global_rows[(tm.ref.idx, r_idx)] = row_ptr
        note_row_map[(tm.ref.idx, r_idx)] = row_ptr

        # 右侧：只写差额公式，原始数字直接在公式里引用「原始数据」
        for cc in report_by_row.get((tm.ref.idx, r_idx), []):
            if cc.period in report_map:
                dc = report_map[cc.period]
                raw_ref = raw.ref(cc.sheet, cc.rows, cc.cols) if raw else None
                parts = []
                for i, (rr, cc0, coef) in enumerate(
                        cc.note_terms or [(cc.note_row, cc.note_col, 1)]):
                    ref = f'{get_column_letter(cc0)}{row_ptr}'
                    if i == 0:
                        parts.append(f'-{ref}' if coef < 0 else ref)
                    else:
                        parts.append(('-' if coef < 0 else '+') + ref)
                expr = ''.join(parts)
                if raw_ref:
                    formula = f'=({expr})-{raw_ref}'
                else:
                    formula = f'=({expr})-{cc.report_value}'
                cell_diff = ws.cell(row_ptr, dc, formula)
                cell_diff.number_format = FMT_AMOUNT
                cell_diff.border = BORDER
                cell_diff.font = F_REPORT
                cell_diff.fill = (FILL_PASS if cc.status == 'pass'
                                  else FILL_FAIL_TARGET if cc.status == 'fail'
                                  else FILL_UNKNOWN)
                stats['report_cells'] += 1

        for tc in tb_by_row.get((tm.ref.idx, r_idx), []):
            if tc.period in tb_map:
                tdcol = tb_map[tc.period]
                raw_ref = raw.ref(tc.tb_sheet, tc.tb_rows, tc.tb_cols) if raw else None
                note_ref = f'{get_column_letter(tc.note_col)}{row_ptr}'
                formula = (f'={note_ref}-{raw_ref}' if raw_ref
                           else f'={note_ref}-{tc.tb_value}')
                cell_diff = ws.cell(row_ptr, tdcol, formula)
                cell_diff.number_format = FMT_AMOUNT
                cell_diff.border = BORDER
                cell_diff.font = F_REPORT
                cell_diff.fill = (FILL_PASS if tc.status == 'pass'
                                  else FILL_FAIL_TARGET if tc.status == 'fail'
                                  else FILL_UNKNOWN)
                stats['tb_cells'] += 1

        row_ptr += 1

        for item in checks_by_row.get(r_idx, []):
            row_ptr = _write_check_row(ws, item, src_to_dst, results_by_key,
                                       row_ptr, stats, global_rows, check_rows)

    return row_ptr, note_row_map


def _cell_ref(term, src_to_dst, global_rows, own_table_idx):
    tid = term.table_idx or own_table_idx
    if tid == own_table_idx:
        row = src_to_dst.get(term.row, term.row)
    else:
        row = global_rows.get((tid, term.row))
        if row is None:
            return None
    return f'{get_column_letter(term.col)}{row}'


def _terms_expr(item, terms, src_to_dst, global_rows, own_table_idx):
    parts = []
    for i, t in enumerate(terms):
        ref = _cell_ref(t, src_to_dst, global_rows, own_table_idx)
        if ref is None:
            return None
        sign = '-' if t.coef < 0 else '+'
        if i == 0:
            parts.append(f'-{ref}' if t.coef < 0 else ref)
        else:
            parts.append(f'{sign}{ref}')
    return ''.join(parts)


def _write_check_row(ws, item, src_to_dst, results_by_key, row_ptr, stats, global_rows,
                     check_rows=None):
    res = results_by_key.get(id(item))
    own_idx = item.table_idx
    target_ref = (f'{get_column_letter(item.target_col)}'
                  f'{src_to_dst.get(item.target_row, item.target_row)}')
    if check_rows is not None:
        check_rows[id(item)] = row_ptr

    label = f'【勾稽校验】{item.label}'
    cell = ws.cell(row_ptr, 1, label)
    cell.font = F_CHECK
    cell.fill = FILL_FAIL_ROW if (res and res.status == 'fail') else FILL_CHECK_ROW

    if item.den:
        num = _terms_expr(item, item.terms, src_to_dst, global_rows, own_idx)
        den = _terms_expr(item, item.den, src_to_dst, global_rows, own_idx)
        formula = (f'={target_ref}-({num})/({den})'
                   if num is not None and den is not None else None)
        number_format = FMT_RATIO
    else:
        expr = _terms_expr(item, item.terms, src_to_dst, global_rows, own_idx)
        formula = f'={target_ref}-({expr})' if expr is not None else None
        number_format = FMT_AMOUNT

    if formula is None:
        target_cell = ws.cell(row_ptr, item.target_col, '（跨表引用不可用）')
        target_cell.fill = FILL_UNKNOWN
        target_cell.font = F_CHECK
        target_cell.border = BORDER
        for c in range(1, item.target_col + 1):
            ws.cell(row_ptr, c).border = BORDER
        stats['skip'] += 1
        return row_ptr + 1

    target_cell = ws.cell(row_ptr, item.target_col, formula)
    target_cell.number_format = number_format
    target_cell.font = F_CHECK
    target_cell.border = BORDER

    if res is None or res.status == 'skip':
        target_cell.fill = FILL_UNKNOWN
        stats['skip'] += 1
    elif res.status == 'unclear':
        target_cell.fill = FILL_UNKNOWN
        for c in range(1, item.target_col + 1):
            ws.cell(row_ptr, c).fill = FILL_CHECK_ROW
        ws.cell(row_ptr, 1).fill = FILL_UNKNOWN
        stats['unclear'] += 1
    elif res.status == 'pass':
        target_cell.fill = FILL_PASS
        stats['pass'] += 1
    else:
        target_cell.fill = FILL_FAIL_TARGET
        for c in range(1, item.target_col + 1):
            ws.cell(row_ptr, c).fill = FILL_FAIL_ROW
        ws.cell(row_ptr, item.target_col).fill = FILL_FAIL_TARGET
        stats['fail'] += 1

    for c in range(1, item.target_col + 1):
        ws.cell(row_ptr, c).border = BORDER
    return row_ptr + 1


# ---------------------------------------------------------------- 原始数据

class RawData:
    """把报表 / TB 的数值贴进 `原始数据` 工作表，并记住每张表的行偏移。

    这样底稿右侧的「报表数」「TB数」可以写成**引用公式**（真链接），
    点开就能看到数据来自 `原始数据` 的哪一格，同时不增加太多工作表。
    """

    SHEET_NAME = '原始数据'

    def __init__(self, ws):
        self.ws = ws
        self.row = 1
        self.offsets = {}

    def add(self, name, src_ws, max_rows=260, max_cols=32):
        ws = self.ws
        ws.cell(self.row, 1, f'【{name}】').font = F_TABLE_TITLE
        ws.cell(self.row, 1).fill = FILL_TABLE_TITLE
        self.row += 1
        self.offsets[name] = self.row - 1
        n_rows = min(src_ws.max_row, max_rows)
        n_cols = min(src_ws.max_column, max_cols)
        for r in range(1, n_rows + 1):
            for c in range(1, n_cols + 1):
                v = src_ws.cell(r, c).value
                if v is None:
                    continue
                cell = ws.cell(self.offsets[name] + r, c, v)
                if isinstance(v, float):
                    cell.number_format = FMT_AMOUNT
        self.row = self.offsets[name] + n_rows + 3

    def ref(self, name, row, col):
        off = self.offsets.get(name)
        if off is None:
            return None
        return f"'{self.SHEET_NAME}'!{get_column_letter(col)}{off + row}"


def _build_raw_data(wb, report, tb_xlsx, tb_sheet_names):
    """建 原始数据 工作表，返回 RawData 句柄。"""
    ws = wb.create_sheet(RawData.SHEET_NAME)
    raw = RawData(ws)
    import openpyxl
    if report is not None and getattr(report, 'path', None):
        try:
            src = openpyxl.load_workbook(report.path, data_only=True)
            for name in src.sheetnames:
                raw.add(name, src[name])
        except Exception:
            pass
    if tb_xlsx:
        try:
            src = openpyxl.load_workbook(tb_xlsx, data_only=True)
            for name in (tb_sheet_names or []):
                if name in src.sheetnames:
                    raw.add(name, src[name], max_rows=120, max_cols=24)
        except Exception:
            pass
    return raw


# ---------------------------------------------------------------- 汇总表

def _block_header(ws, row, title, headers, widths=None):
    c = ws.cell(row, 1, title)
    c.font = Font(name='微软雅黑', size=11, bold=True, color='1F497D')
    row += 1
    for i, h in enumerate(headers, 1):
        cell = ws.cell(row, i, h)
        cell.font = F_TITLE
        cell.fill = FILL_TITLE
        cell.alignment = Alignment(horizontal='center')
        cell.border = BORDER
    return row + 1



def _report_item_block(ws, r, report, cross, note_no_checks, note_rows, stats):
    """二、报表项目逐项勾稽：每个报表项目一行，写清金额、是否勾稽、是否一致。

    汇总表允许贴金额（看得清楚）；勾稽底稿不贴，只留公式。
    """
    try:
        import openpyxl
        book = openpyxl.load_workbook(report.path, data_only=True)
    except Exception:
        book = None

    def cell_value(sheet_name, row, col):
        if book is None or sheet_name not in book.sheetnames:
            return None
        v = book[sheet_name].cell(row, col).value
        if isinstance(v, bool):
            return None
        if isinstance(v, (int, float)):
            return float(v)
        t = str(v or '').strip().replace(',', '')
        if t in ('', '-', '--', '—', '——', '–', '/'):
            return 0.0
        try:
            return float(t)
        except ValueError:
            return None

    by_row = {}
    for cc in (cross or []):
        by_row.setdefault((cc.sheet, cc.rows), []).append(cc)

    import report_reader as _rr
    note_no = {}
    for chk in (note_no_checks or []):
        note_no[(chk.sheet, _rr.normalize_name(chk.item_name))] = chk

    # 先扫一遍，决定要展示几期（最多两期：期末/期初）
    periods = []
    for sheet in (report.sheets if report else []):
        for _col, d in sheet.period_columns():
            if d not in periods:
                periods.append(d)
    periods = sorted(periods, reverse=True)[:2]

    headers = ['序号', '报表项目', '附注编号', '附注科目']
    for d in periods:
        headers += [f'报表数 {d}', f'附注数 {d}', f'差额 {d}']
    headers += ['是否勾稽', '是否一致', '定位']
    col_flag = 4 + len(periods) * 3 + 1
    col_state = col_flag + 1
    col_link = col_state + 1

    r = _block_header(ws, r, '二、报表项目逐项勾稽（每个报表项目一行）', headers)

    n = tot = ok = bad = seq = 0
    for sheet in (report.sheets if report else []):
        if not sheet.period_columns():
            continue
        block = sheet.blocks or {}
        for name, (row, item_col) in sheet.items.items():
            if not name or name.endswith((':', '：')):
                continue
            if any(w in name for w in ('合计', '总计')):
                continue
            # 营业总收入/营业总成本这类是"总计"概念，不是报表项目；
            # 其他综合收益的子项同理（见 report_numbering.NO_NOTE_ITEMS）
            try:
                import report_numbering as _rn
                if _rn.is_no_note_item(name):
                    continue
            except Exception:
                pass
            tot += 1
            val_cols = set((block.get(item_col) or {}).keys())
            checks = [c for c in by_row.get((sheet.name, row), [])
                      if c.cols in val_cols]
            by_period = {c.period: c for c in checks}
            nc = note_no.get((sheet.name, _rr.normalize_name(name)))

            vals = [0, name,
                    (nc.actual if nc else ''),
                    (checks[0].sub_item if checks else '')]
            any_amount = False
            for d in periods:
                vcol = (block.get(item_col) or {}).get(d)
                # 期间 -> 列号 反查
                if vcol is None:
                    for cc0, dd in (block.get(item_col) or {}).items():
                        if dd == d:
                            vcol = cc0
                            break
                rep_val = cell_value(sheet.name, row, vcol) if vcol else None
                chk = by_period.get(d)
                note_val = chk.note_value if chk else None
                if rep_val or note_val:
                    any_amount = True
                vals += [rep_val, note_val, (chk.diff if chk else None)]

            if checks:
                n += 1
                if any(c.status == 'fail' for c in checks):
                    flag, state = '已勾稽', '有差异'
                    bad += 1
                elif any(c.status == 'unmatched' for c in checks):
                    flag, state = '已勾稽', '待确认'
                else:
                    flag, state = '已勾稽', '一致'
                    ok += 1
            else:
                flag, state = ('未勾稽', '') if any_amount else ('—（无金额）', '')
            if not any_amount and not checks:
                # 报表和附注都是 0/空：不显示这一行
                continue
            seq += 1
            vals[0] = seq
            vals += [flag, state, '']

            for i, v in enumerate(vals, 1):
                cell = ws.cell(r, i, v)
                cell.font = F_BODY
                cell.border = BORDER
                if 5 <= i < col_flag and isinstance(v, (int, float)):
                    cell.number_format = FMT_AMOUNT
            color = {'一致': FILL_PASS, '有差异': FILL_FAIL_TARGET,
                     '待确认': FILL_UNKNOWN}.get(state)
            if color:
                ws.cell(r, col_state).fill = color
            if flag in ('未勾稽',):
                ws.cell(r, col_flag).fill = FILL_UNKNOWN
            if checks:
                target = (note_rows or {}).get((checks[0].note_table,
                                                checks[0].note_row))
                if target:
                    cell = ws.cell(r, col_link, f'勾稽底稿!A{target}')
                    cell.font = Font(name='微软雅黑', size=9, color='0563C1',
                                     underline='single')
                    from openpyxl.worksheet.hyperlink import Hyperlink
                    cell.hyperlink = Hyperlink(
                        ref=cell.coordinate, location=f"'勾稽底稿'!A{target}")
            r += 1

    ws.cell(r, 1, f'小计：报表项目 {seq} 个，已勾稽 {n} 个'
                  f'（一致 {ok}，有差异 {bad}）').font = F_HEADER
    r += 2
    return r


def _write_summary_sheet(ws, stats, results, model, tables, cross_count,
                         tb_count, nn_count, nx_count, disclosure_count,
                         source_note='', check_rows=None, note_rows=None,
                         cross=None, tb_checks=None, note_cross=None,
                         report=None, note_no_checks=None):
    ws.cell(1, 1, '勾稽汇总表').font = Font(name='微软雅黑', size=14, bold=True, color='1F497D')
    r = 3

    # 一、总览
    ws.cell(r, 1, '一、勾稽总览').font = F_HEADER
    r += 1
    rows = [('附注文件', model.path),
            ('本次采用的底稿', source_note or '原始文件'),
            ('主要项目附注章', len(model.chapters)),
            ('子项目合计', len(model.sub_items)),
            ('表格总数', len(tables))]
    for k, v in rows:
        ws.cell(r, 1, k).font = F_BODY
        ws.cell(r, 2, v).font = F_BODY
        r += 1

    def group(name, total, ok, bad, other=None, other_label='待确认'):
        nonlocal r
        ws.cell(r, 1, name).font = F_BODY
        ws.cell(r, 2, total).font = F_BODY
        ws.cell(r, 3, f'平 {ok}').font = F_BODY
        bad_cell = ws.cell(r, 4, f'有差异 {bad}')
        bad_cell.font = Font(name='微软雅黑', size=9, bold=True,
                             color='9C0006' if bad else '000000')
        if other is not None:
            ws.cell(r, 5, f'{other_label} {other}').font = F_BODY
        r += 1

    r += 1
    for i, h in enumerate(['核对类型', '条数', '结果', '', ''], 1):
        c = ws.cell(r, i, h)
        if h:
            c.font = F_HEADER
    r += 1
    group('附注内部勾稽',
          stats['pass'] + stats['fail'] + stats['skip'] + stats['unclear'],
          stats['pass'], stats['fail'], stats['skip'], '无法计算')
    if stats['unclear']:
        ws.cell(r, 1, '  其中结构存疑（黄色）').font = F_BODY
        ws.cell(r, 2, stats['unclear']).font = F_BODY
        ws.cell(r, 3, '认不准结构，请人工确认这些表').font = Font(
            name='微软雅黑', size=9, color='B25E00')
        r += 1
    group('报表勾稽', stats['cross_pass'] + stats['cross_fail'] + stats['cross_unmatched'],
          stats['cross_pass'], stats['cross_fail'], stats['cross_unmatched'])
    if tb_count:
        group('TB现金流核对', stats['tb_pass'] + stats['tb_fail'] + stats['tb_unmatched'],
              stats['tb_pass'], stats['tb_fail'], stats['tb_unmatched'])
    else:
        ws.cell(r, 1, 'TB现金流核对').font = F_BODY
        ws.cell(r, 2, '未导入 TB').font = F_BODY
        ws.cell(r, 3, '现金流量表子项未与 TB 核对').font = Font(
            name='微软雅黑', size=9, color='B25E00')
        r += 1
    group('附注号勾稽', stats['nn_pass'] + stats['nn_fail'] + stats['nn_unmatched'],
          stats['nn_pass'], stats['nn_fail'], stats['nn_unmatched'])
    group('附注内跨表', stats['nx_pass'] + stats['nx_fail'] + stats['nx_unmatched'],
          stats['nx_pass'], stats['nx_fail'], stats['nx_unmatched'])
    ws.cell(r, 1, '未披露候选（人工判断）').font = F_BODY
    ws.cell(r, 2, disclosure_count).font = F_BODY
    r += 2

    # 二、报表项目逐项勾稽（参考移植前看板：先逐项列清）
    r = _report_item_block(ws, r, report, cross, note_no_checks, note_rows, stats)

    # 三、差异明细（带跳转链接，点一下跳到底稿对应位置）
    r = _block_header(ws, r, '三、差异明细（点「定位」可跳转到勾稽底稿对应位置）',
                      ['序号', '类型', '子项目 / 说明', '期间', '期望值', '实际值',
                       '差额', '定位'])
    from openpyxl.worksheet.hyperlink import Hyperlink

    def add_link(row, col, target_row):
        if not target_row:
            return
        cell = ws.cell(row, col, f'勾稽底稿!A{target_row}')
        cell.font = Font(name='微软雅黑', size=9, color='0563C1', underline='single')
        cell.hyperlink = Hyperlink(ref=cell.coordinate,
                                   location=f"'勾稽底稿'!A{target_row}")

    n = 0
    # 内部勾稽差异
    for res in [x for x in results if x.status == 'fail']:
        it = res.item
        n += 1
        period = str(it.period) if it.period else (it.period_key or '')
        vals = [n, '附注内部勾稽', f'{it.sub_item}：{it.label}', period,
                res.expected, res.actual, res.diff, '']
        for i, v in enumerate(vals, 1):
            cell = ws.cell(r, i, v)
            cell.font = F_BODY
            cell.border = BORDER
            if i in (5, 6, 7) and isinstance(v, (int, float)):
                cell.number_format = ('0.00%;[Red]-0.00%' if getattr(it, 'is_ratio', False)
                                      else FMT_AMOUNT)
        ws.cell(r, 7).fill = FILL_FAIL_TARGET
        add_link(r, 8, (check_rows or {}).get(id(it)))
        r += 1

    # 报表勾稽差异
    for cc in [x for x in (cross or []) if x.status == 'fail']:
        n += 1
        vals = [n, '报表勾稽', f'{cc.sub_item}（{cc.sheet} {cc.rows}行）',
                str(cc.period) if cc.period else '', cc.report_value, cc.note_value,
                cc.diff, '']
        for i, v in enumerate(vals, 1):
            cell = ws.cell(r, i, v)
            cell.font = F_BODY
            cell.border = BORDER
            if i in (5, 6, 7) and isinstance(v, (int, float)):
                cell.number_format = FMT_AMOUNT
        ws.cell(r, 7).fill = FILL_FAIL_TARGET
        add_link(r, 8, (note_rows or {}).get((cc.note_table, cc.note_row)))
        r += 1

    # TB 核对差异
    for tc in [x for x in (tb_checks or []) if x.status == 'fail']:
        n += 1
        vals = [n, 'TB现金流核对', f'{tc.sub_item}：{tc.label}',
                str(tc.period) if tc.period else '', tc.tb_value, tc.note_value,
                tc.diff, '']
        for i, v in enumerate(vals, 1):
            cell = ws.cell(r, i, v)
            cell.font = F_BODY
            cell.border = BORDER
            if i in (5, 6, 7) and isinstance(v, (int, float)):
                cell.number_format = FMT_AMOUNT
        ws.cell(r, 7).fill = FILL_FAIL_TARGET
        add_link(r, 8, (note_rows or {}).get((tc.table_idx, tc.note_row)))
        r += 1

    # 附注内跨表差异
    for xc in [x for x in (note_cross or []) if x.status == 'fail']:
        n += 1
        vals = [n, '附注内跨表', f'{xc.sub_a}：{xc.note}',
                str(xc.period) if xc.period else '', xc.value_a, xc.value_b,
                xc.diff, '']
        for i, v in enumerate(vals, 1):
            cell = ws.cell(r, i, v)
            cell.font = F_BODY
            cell.border = BORDER
            if i in (5, 6, 7) and isinstance(v, (int, float)):
                cell.number_format = FMT_AMOUNT
        ws.cell(r, 7).fill = FILL_FAIL_TARGET
        c = ws.cell(r, 8, '见「完整性」第三部分')
        c.font = Font(name='微软雅黑', size=9, color='666666')
        r += 1

    if n == 0:
        ws.cell(r, 1, '（全部通过，无差异）').font = F_BODY
        r += 1
    r += 1

    # 三、按规则 / 表类型分布
    by_rule = Counter(x.item.rule_id for x in results)
    rule_name = {'V-SUM': '合计行 = 上方明细之和',
                 'V-GRP': '「其中：」父行 = 子项之和',
                 'H-SUM': '合计列 = 其余列之和',
                 'R-NET': '账面价值 = 账面余额 − 准备',
                 'R-NBV': '账面价值 = 原值 − 折旧 − 减值',
                 'R-MOVE': '期末 = 期初 + 本期增加 − 本期减少',
                 'R-PCT': '比例重算'}
    ws.cell(r, 1, '四、规则与表类型分布').font = F_HEADER
    r += 1
    for k, v in sorted(by_rule.items()):
        ws.cell(r, 1, k).font = F_BODY
        ws.cell(r, 2, f'{rule_name.get(k, "")}').font = F_BODY
        ws.cell(r, 3, f'{v} 条').font = F_BODY
        r += 1
    try:
        import table_types
        for name, n in table_types.summarize(tables).items():
            ws.cell(r, 1, '表类型').font = F_BODY
            ws.cell(r, 2, name).font = F_BODY
            ws.cell(r, 3, f'{n} 张').font = F_BODY
            r += 1
    except Exception:
        pass

    for col, w in (('A', 26), ('B', 46), ('C', 18), ('D', 18), ('E', 16),
                   ('F', 46), ('G', 12), ('H', 16), ('I', 16), ('J', 14)):
        ws.column_dimensions[col].width = w


# ---------------------------------------------------------------- 完整性

def _write_integrity_sheet(ws, disclosure, note_no_checks, note_cross, anomalies=None):
    ws.cell(1, 1, '完整性检查').font = Font(name='微软雅黑', size=14, bold=True, color='1F497D')
    r = 3

    # 一、未披露
    ws.cell(r, 1, '一、报表项目未在附注中披露（供人工判断是漏披露还是无需披露）').font = F_HEADER
    r += 1
    periods = sorted({d for it in (disclosure or []) for d in it.amounts}, reverse=True)
    headers = ['序号', '报表', '行号', '报表科目'] + [str(d) for d in periods] + ['建议']
    for i, h in enumerate(headers, 1):
        c = ws.cell(r, i, h)
        c.font = F_TITLE
        c.fill = FILL_TITLE
        c.alignment = Alignment(horizontal='center')
        c.border = BORDER
    r += 1
    for n, it in enumerate(disclosure or [], 1):
        vals = [n, it.sheet, it.row, it.item_name] + \
               [it.amounts.get(d) for d in periods] + ['待人工判断']
        for i, v in enumerate(vals, 1):
            cell = ws.cell(r, i, v)
            cell.font = F_BODY
            cell.border = BORDER
            if 5 <= i <= 4 + len(periods):
                cell.number_format = FMT_AMOUNT
        ws.cell(r, len(vals)).fill = FILL_UNKNOWN
        r += 1
    if not disclosure:
        ws.cell(r, 1, '（无）').font = F_BODY
        r += 1
    r += 2

    # 二、附注号勾稽
    ws.cell(r, 1, '二、报表「附注」列 ↔ 附注实际编号').font = F_HEADER
    r += 1
    for i, h in enumerate(['序号', '工作表', '行号', '科目', '报表附注号',
                           '附注实际编号', '状态', '说明'], 1):
        c = ws.cell(r, i, h)
        c.font = F_TITLE
        c.fill = FILL_TITLE
        c.alignment = Alignment(horizontal='center')
        c.border = BORDER
    r += 1
    for n, chk in enumerate(note_no_checks or [], 1):
        vals = [n, chk.sheet, chk.row, chk.item_name, chk.actual or '（空）',
                chk.expected, '', chk.note]
        for i, v in enumerate(vals, 1):
            cell = ws.cell(r, i, v)
            cell.font = F_BODY
            cell.border = BORDER
        if chk.status == 'pass':
            ws.cell(r, 7, '一致')
            ws.cell(r, 7).fill = FILL_PASS
        elif chk.status == 'fail':
            ws.cell(r, 7, '不一致')
            for c in range(1, 8):
                ws.cell(r, c).fill = FILL_FAIL_ROW
            ws.cell(r, 7).font = Font(name='微软雅黑', size=9, color='9C0006')
        else:
            ws.cell(r, 7, '待确认')
            ws.cell(r, 7).fill = FILL_UNKNOWN
        r += 1
    if not note_no_checks:
        ws.cell(r, 1, '（无）').font = F_BODY
        r += 1
    r += 2

    # 三、附注内跨表
    ws.cell(r, 1, '三、附注内跨表 / 上下文相关项目').font = F_HEADER
    r += 1
    for i, h in enumerate(['序号', '类型', '期间', '子项目A', '表A', '金额A',
                           '子项目B', '表B', '金额B', '差额', '状态', '说明'], 1):
        c = ws.cell(r, i, h)
        c.font = F_TITLE
        c.fill = FILL_TITLE
        c.alignment = Alignment(horizontal='center')
        c.border = BORDER
    r += 1
    for n, chk in enumerate(note_cross or [], 1):
        vals = [n, chk.kind, str(chk.period) if chk.period else '', chk.sub_a,
                f'表#{chk.table_a}', chk.value_a, chk.sub_b, f'表#{chk.table_b}',
                chk.value_b, chk.diff, '', chk.note]
        for i, v in enumerate(vals, 1):
            cell = ws.cell(r, i, v)
            cell.font = F_BODY
            cell.border = BORDER
            if i in (6, 9, 10):
                cell.number_format = FMT_AMOUNT
        if chk.status == 'pass':
            ws.cell(r, 11, '平')
            ws.cell(r, 10).fill = FILL_PASS
        elif chk.status == 'fail':
            ws.cell(r, 11, '有差异')
            for c in range(1, 11):
                ws.cell(r, c).fill = FILL_FAIL_ROW
            ws.cell(r, 10).fill = FILL_FAIL_TARGET
            ws.cell(r, 11).font = Font(name='微软雅黑', size=9, color='9C0006')
        else:
            ws.cell(r, 11, '待确认')
            ws.cell(r, 11).fill = FILL_UNKNOWN
        r += 1
    if not note_cross:
        ws.cell(r, 1, '（无）').font = F_BODY
        r += 1
    r += 2

    # 四、数值格式异常
    ws.cell(r, 1, '四、数值格式异常（尾随负号，数值已按负数参与勾稽，请把「-」删掉）').font = F_HEADER
    r += 1
    for i, h in enumerate(['序号', '子项目', '表号', '表内行', '表内列', '原值', '处理建议'], 1):
        c = ws.cell(r, i, h)
        c.font = F_TITLE
        c.fill = FILL_TITLE
        c.alignment = Alignment(horizontal='center')
        c.border = BORDER
    r += 1
    for n, a in enumerate(anomalies or [], 1):
        vals = [n, a['sub_item'], f"表#{a['table']}", a['row'], a['col'],
                a['text'], '删除尾随「-」，改为标准负数或 0']
        for i, v in enumerate(vals, 1):
            cell = ws.cell(r, i, v)
            cell.font = F_BODY
            cell.border = BORDER
        ws.cell(r, 6).fill = FILL_UNKNOWN
        r += 1
    if not anomalies:
        ws.cell(r, 1, '（无）').font = F_BODY
        r += 1

    for col, w in (('A', 8), ('B', 14), ('C', 14), ('D', 34), ('E', 16),
                   ('F', 16), ('G', 22), ('H', 10), ('I', 16), ('J', 14),
                   ('K', 10), ('L', 40)):
        ws.column_dimensions[col].width = w


# ---------------------------------------------------------------- 入口

def write_workbook(out_path, model, tables, checks, results, cross=None,
                   report=None, tb_checks=None, tb_xlsx=None,
                   note_no_checks=None, disclosure=None, note_cross=None,
                   source_note='', log=print):
    """生成成果表：勾稽底稿 / 汇总表 / 完整性 三张。"""
    wb = Workbook()
    ws_sum = wb.active          # 汇总表放第一张
    ws_sum.title = '汇总表'
    ws = wb.create_sheet('勾稽底稿')   # 主体第二张

    # 原始数据（报表 + TB）先贴好，底稿右侧才能写成引用
    tb_sheet_names = sorted({t.tb_sheet for t in (tb_checks or []) if t.tb_sheet})
    raw = _build_raw_data(wb, report, tb_xlsx, tb_sheet_names)

    ws.cell(1, 1, '财务报表附注勾稽底稿').font = Font(name='微软雅黑', size=14, bold=True,
                                              color='1F497D')
    ws.cell(2, 1, '绿色=已平（差额≤0.01）　深红=有差异　淡红底=该条有差异　'
                  '黄色=结构存疑/无法确认').font = Font(name='微软雅黑', size=9, color='666666')

    stats = Counter({'pass': 0, 'fail': 0, 'skip': 0, 'unclear_tables': 0,
                     'cross_pass': 0, 'cross_fail': 0, 'cross_unmatched': 0,
                     'tb_pass': 0, 'tb_fail': 0, 'tb_unmatched': 0,
                     'nn_pass': 0, 'nn_fail': 0, 'nn_unmatched': 0,
                     'nx_pass': 0, 'nx_fail': 0, 'nx_unmatched': 0,
                     'missing_count': 0, 'report_cells': 0, 'tb_cells': 0,
                     'unclear': 0})

    # 先统计各核对类型的结果（用于汇总表）
    for cc in (cross or []):
        if cc.status == 'pass':
            stats['cross_pass'] += 1
        elif cc.status == 'fail':
            stats['cross_fail'] += 1
        else:
            stats['cross_unmatched'] += 1
    for tc in (tb_checks or []):
        if tc.status == 'pass':
            stats['tb_pass'] += 1
        elif tc.status == 'fail':
            stats['tb_fail'] += 1
        else:
            stats['tb_unmatched'] += 1
    for nc in (note_no_checks or []):
        if nc.status == 'pass':
            stats['nn_pass'] += 1
        elif nc.status == 'fail':
            stats['nn_fail'] += 1
        else:
            stats['nn_unmatched'] += 1
    for xc in (note_cross or []):
        if xc.status == 'pass':
            stats['nx_pass'] += 1
        elif xc.status == 'fail':
            stats['nx_fail'] += 1
        else:
            stats['nx_unmatched'] += 1
    stats['missing_count'] = len(disclosure or [])

    periods = _period_slots(cross, tb_checks)
    report_map, tb_map = _col_map(periods)

    # 右侧表头
    if periods:
        _write_right_headers(ws, 3, periods, report_map, tb_map, bool(tb_checks))

    results_by_key = {id(r.item): r for r in results}
    checks_by_table = {}
    for item in checks:
        checks_by_table.setdefault((item.chapter, item.sub_item, item.table_idx),
                                   {}).setdefault(item.target_row, []).append(item)

    report_by_row = {}
    for cc in (cross or []):
        report_by_row.setdefault((cc.note_table, cc.note_row), []).append(cc)
    tb_by_row = {}
    for tc in (tb_checks or []):
        tb_by_row.setdefault((tc.table_idx, tc.note_row), []).append(tc)

    global_rows = {}
    note_rows = {}
    check_rows = {}
    anomalies = []
    row_ptr = 5
    for tm in tables:
        key = (tm.chapter, tm.sub_item, tm.ref.idx)
        row_ptr, nrm = _write_table(ws, tm, checks_by_table.get(key, {}),
                                    results_by_key, row_ptr, stats, global_rows,
                                    report_by_row, tb_by_row, report_map, tb_map,
                                    raw, check_rows, anomalies)
        note_rows.update(nrm)
        row_ptr += 1

    ws.column_dimensions['A'].width = 46
    for c in range(2, 11):
        ws.column_dimensions[get_column_letter(c)].width = 16
    for p in periods:
        ws.column_dimensions[get_column_letter(report_map[p])].width = 18
        if tb_checks:
            ws.column_dimensions[get_column_letter(tb_map[p])].width = 18
    ws.freeze_panes = 'A5'

    _write_summary_sheet(ws_sum, stats, results, model, tables,
                         len(cross or []), len(tb_checks or []),
                         len(note_no_checks or []), len(note_cross or []),
                         len(disclosure or []), source_note=source_note,
                         check_rows=check_rows, note_rows=note_rows,
                         cross=cross, tb_checks=tb_checks, note_cross=note_cross,
                         report=report, note_no_checks=note_no_checks)
    _write_integrity_sheet(wb.create_sheet('完整性'), disclosure,
                           note_no_checks, note_cross, anomalies)

    # 工作表顺序：汇总表 → 勾稽底稿 → 完整性 → 原始数据
    order = ['汇总表', '勾稽底稿', '完整性', RawData.SHEET_NAME]
    wb._sheets.sort(key=lambda sh: order.index(sh.title)
                    if sh.title in order else len(order))
    wb.save(out_path)
    log(f'[输出] 成果表已生成：{out_path}')
    log(f'[输出] 附注内部勾稽 {len(results)} 条：通过 {stats["pass"]}，'
        f'有差异 {stats["fail"]}，结构存疑 {stats["unclear"]}，'
        f'无法计算 {stats["skip"]}')
    if cross:
        log(f'[输出] 报表勾稽 {len(cross)} 条：平 {stats["cross_pass"]}，'
            f'有差异 {stats["cross_fail"]}，待确认 {stats["cross_unmatched"]}')
    if tb_checks:
        log(f'[输出] TB现金流核对 {len(tb_checks)} 条：平 {stats["tb_pass"]}，'
            f'有差异 {stats["tb_fail"]}，待确认 {stats["tb_unmatched"]}')
    if note_no_checks:
        log(f'[输出] 附注号勾稽 {len(note_no_checks)} 条：一致 {stats["nn_pass"]}，'
            f'不一致 {stats["nn_fail"]}，待确认 {stats["nn_unmatched"]}')
    if note_cross:
        log(f'[输出] 附注内跨表 {len(note_cross)} 条：平 {stats["nx_pass"]}，'
            f'有差异 {stats["nx_fail"]}，待确认 {stats["nx_unmatched"]}')
    if disclosure is not None:
        log(f'[输出] 未披露汇总 {len(disclosure)} 条（供人工判断）')
    if anomalies:
        log(f'[输出] 数值格式异常 {len(anomalies)} 处（尾随负号，已按负数参与勾稽，'
            f'建议删除「-」）')
    return out_path, dict(stats)
