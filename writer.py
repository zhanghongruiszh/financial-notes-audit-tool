# -*- coding: utf-8 -*-
"""
writer.py — 把勾稽结果写成 Excel 成果表。

成果表包含
----------
1. `勾稽底稿`：逐表还原附注内容，**在每个勾稽点的目标行下方**插入一行
   `【勾稽校验】…`，单元格里写活公式（点进去就能看见引用了哪几格）。
      · 差额为 0（≤1 分钱）→ 绿色
      · 差额非 0          → 差额格深红 + 整行淡红底
      · 结构没认出来的表   → 标题行标黄，并写一句原因
2. `差异清单`：把所有有差异的勾稽逐条列出，便于跟进
3. `勾稽总览`：条数统计

设计取舍
--------
不做"列位一一对应"的还原（Word 合并单元格千奇百怪），而是保证
**「一个科目下方紧跟着它的校验行」**，公式直接指向参与运算的单元格，
这样点公式就能立刻定位问题。
"""

import datetime
from collections import Counter

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

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

FILL_PASS = PatternFill('solid', start_color='E2EFDA', end_color='E2EFDA')
FILL_FAIL_TARGET = PatternFill('solid', start_color='FFC7CE', end_color='FFC7CE')
FILL_FAIL_ROW = PatternFill('solid', start_color='FDE2E2', end_color='FDE2E2')
FILL_UNKNOWN = PatternFill('solid', start_color='FFF2CC', end_color='FFF2CC')
FILL_CHECK_ROW = PatternFill('solid', start_color='F7F9FC', end_color='F7F9FC')

THIN = Side(style='thin', color='D9D9D9')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FMT_AMOUNT = '#,##0.00;[Red](#,##0.00);"-"'
FMT_RATIO = '0.00%;[Red]-0.00%;"-"'
FMT_DATE = 'yyyy-mm-dd'


def _write_value(ws, row, col, value):
    """写一个单元格并按类型设格式。"""
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


# ---------------------------------------------------------------- 底稿

def _write_table(ws, tm, checks_by_row, results_by_key, row_ptr, stats):
    """写一张表，并在勾稽点下方插校验行。返回下一可用行号。"""
    grid = tm.ref.grid

    # 表标题
    ws.cell(row_ptr, 1, f'{tm.chapter} ｜ {tm.sub_item} ｜ 表#{tm.ref.idx} '
                        f'（{tm.ref.rows}行×{tm.ref.cols}列）'
                        f'｜ {tm.table_type_name or "未分类"}')
    ws.cell(row_ptr, 1).font = F_TABLE_TITLE
    ws.cell(row_ptr, 1).fill = FILL_TABLE_TITLE
    for c in range(2, max(3, tm.ref.cols) + 1):
        ws.cell(row_ptr, c).fill = FILL_TABLE_TITLE
    row_ptr += 1

    # 结构没认出来 → 标黄 + 说明；期间待确认只作提示，不整表标黄
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
    for r_idx, row in enumerate(grid, 1):
        for c_idx, cellinfo in enumerate(row, 1):
            if cellinfo['m'] != ' ':
                continue
            text = cellinfo['t']
            if not text:
                continue
            import docx_ir
            value = docx_ir.parse_cell_value(text)
            _write_value(ws, row_ptr, c_idx, value)
            if _column_is_ratio(tm, c_idx - 1) and isinstance(value, float):
                ws.cell(row_ptr, c_idx).number_format = FMT_RATIO
        # 表头行加底色
        kind = next((r.kind for r in tm.rows if r.row == r_idx), None)
        if kind == 'header':
            for c_idx in range(1, tm.ref.cols + 1):
                cell = ws.cell(row_ptr, c_idx)
                cell.font = F_HEADER
                cell.fill = FILL_HEADER
        src_to_dst[r_idx] = row_ptr
        row_ptr += 1

        # 这一行是勾稽目标 → 紧跟着插校验行
        for item in checks_by_row.get(r_idx, []):
            row_ptr = _write_check_row(ws, item, src_to_dst, results_by_key, row_ptr, stats)

    return row_ptr


def _term_expr(item, src_to_dst, terms=None):
    """把运算项翻译成 Excel 公式里的单元格表达式。"""
    terms = item.terms if terms is None else terms
    parts = []
    for i, t in enumerate(terms):
        ref = f'{get_column_letter(t.col)}{src_to_dst.get(t.row, t.row)}'
        sign = '-' if t.coef < 0 else '+'
        if i == 0:
            parts.append(f'-{ref}' if t.coef < 0 else ref)
        else:
            parts.append(f'{sign}{ref}')
    return ''.join(parts)


def _write_check_row(ws, item, src_to_dst, results_by_key, row_ptr, stats):
    """写一行【勾稽校验】。"""
    res = results_by_key.get(id(item))
    target_ref = (f'{get_column_letter(item.target_col)}'
                  f'{src_to_dst.get(item.target_row, item.target_row)}')

    label = f'【勾稽校验】{item.label}'
    cell = ws.cell(row_ptr, 1, label)
    cell.font = F_CHECK
    cell.fill = FILL_FAIL_ROW if (res and res.status == 'fail') else FILL_CHECK_ROW

    if item.den:
        num = _term_expr(item, src_to_dst, item.terms)
        den = _term_expr(item, src_to_dst, item.den)
        formula = f'={target_ref}-({num})/({den})'
        number_format = '0.00%;[Red]-0.00%;"-"'
    else:
        formula = f'={target_ref}-({_term_expr(item, src_to_dst)})'
        number_format = FMT_AMOUNT

    target_cell = ws.cell(row_ptr, item.target_col, formula)
    target_cell.number_format = number_format
    target_cell.font = F_CHECK
    target_cell.border = BORDER

    if res is None or res.status == 'skip':
        target_cell.fill = FILL_UNKNOWN
        stats['skip'] += 1
    elif res.status == 'pass':
        target_cell.fill = FILL_PASS
        stats['pass'] += 1
    else:
        target_cell.fill = FILL_FAIL_TARGET      # 差额格深红
        for c in range(1, item.target_col + 1):
            ws.cell(row_ptr, c).fill = FILL_FAIL_ROW   # 整行淡红底
        ws.cell(row_ptr, item.target_col).fill = FILL_FAIL_TARGET
        stats['fail'] += 1

    for c in range(1, item.target_col + 1):
        ws.cell(row_ptr, c).border = BORDER
    return row_ptr + 1


# ---------------------------------------------------------------- 差异清单

def _write_diff_sheet(ws, results):
    headers = ['序号', '章', '子项目', '表号', '规则', '校验说明', '期间',
               '期望值', '实际值', '差额']
    for i, h in enumerate(headers, 1):
        c = ws.cell(1, i, h)
        c.font = F_TITLE
        c.fill = FILL_TITLE
        c.alignment = Alignment(horizontal='center')

    fails = [r for r in results if r.status == 'fail']
    for n, res in enumerate(fails, 1):
        it = res.item
        period = ''
        if it.period:
            period = str(it.period)
        elif it.period_key:
            period = it.period_key
        vals = [n, it.chapter, it.sub_item, it.table_idx, it.rule_id, it.label,
                period, res.expected, res.actual, res.diff]
        for i, v in enumerate(vals, 1):
            cell = ws.cell(n + 1, i, v)
            cell.font = F_BODY
            cell.border = BORDER
            if i in (8, 9, 10) and isinstance(v, (int, float)):
                cell.number_format = FMT_AMOUNT
        ws.cell(n + 1, 10).fill = FILL_FAIL_TARGET

    widths = [6, 22, 18, 7, 9, 46, 12, 15, 15, 13]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'


# ---------------------------------------------------------------- 总览

def _write_summary_sheet(ws, stats, results, model, tables, source_note=''):
    rows = [
        ('附注文件', model.path),
        ('主要项目附注章', len(model.chapters)),
        ('子项目合计', len(model.sub_items)),
        ('表格总数', len(tables)),
        ('附注内部勾稽', len(results)),
        ('  通过', stats['pass']),
        ('  有差异', stats['fail']),
        ('  无法计算', stats['skip']),
        ('结构存疑的表', stats['unclear_tables']),
    ]
    if source_note:
        rows.insert(1, ('本次采用的底稿', source_note))
    if stats.get('cross_pass') or stats.get('cross_fail') or stats.get('cross_unmatched'):
        rows.extend([
            ('报表勾稽', stats['cross_pass'] + stats['cross_fail'] + stats['cross_unmatched']),
            ('  平', stats['cross_pass']),
            ('  有差异', stats['cross_fail']),
            ('  待确认', stats['cross_unmatched']),
        ])
    if stats.get('tb_pass') or stats.get('tb_fail') or stats.get('tb_unmatched'):
        rows.extend([
            ('TB现金流核对', stats['tb_pass'] + stats['tb_fail'] + stats['tb_unmatched']),
            ('  平', stats['tb_pass']),
            ('  有差异', stats['tb_fail']),
            ('  待确认', stats['tb_unmatched']),
        ])
    if stats.get('nn_pass') or stats.get('nn_fail') or stats.get('nn_unmatched'):
        rows.extend([
            ('附注号勾稽', stats['nn_pass'] + stats['nn_fail'] + stats['nn_unmatched']),
            ('  一致', stats['nn_pass']),
            ('  不一致', stats['nn_fail']),
            ('  待确认', stats['nn_unmatched']),
        ])
    if stats.get('nx_pass') or stats.get('nx_fail') or stats.get('nx_unmatched'):
        rows.extend([
            ('附注内跨表', stats['nx_pass'] + stats['nx_fail'] + stats['nx_unmatched']),
            ('  平', stats['nx_pass']),
            ('  有差异', stats['nx_fail']),
            ('  待确认', stats['nx_unmatched']),
        ])
    if stats.get('missing_count'):
        rows.append(('未披露候选（人工判断）', stats['missing_count']))
    ws.cell(1, 1, '勾稽总览').font = Font(name='微软雅黑', size=14, bold=True, color='1F497D')
    for i, (k, v) in enumerate(rows, 3):
        ws.cell(i, 1, k).font = F_HEADER
        c = ws.cell(i, 2, v)
        c.font = F_BODY
        if k.strip() == '有差异' and isinstance(v, int) and v > 0:
            c.font = Font(name='微软雅黑', size=9, bold=True, color='9C0006')
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 70

    # 按规则统计
    row = len(rows) + 5
    ws.cell(row, 1, '按规则分布').font = F_HEADER
    by_rule = Counter(r.item.rule_id for r in results)
    rule_name = {'V-SUM': '合计行 = 上方明细之和',
                 'V-GRP': '「其中：」父行 = 子项之和',
                 'H-SUM': '合计列 = 其余列之和',
                 'R-NET': '账面价值 = 账面余额 − 准备',
                 'R-NBV': '账面价值 = 原值 − 折旧 − 减值',
                 'R-MOVE': '期末 = 期初 + 本期增加 − 本期减少',
                 'X-NOTE': '附注合计 vs 报表科目'}
    for i, (k, v) in enumerate(sorted(by_rule.items()), row + 1):
        ws.cell(i, 1, k).font = F_BODY
        ws.cell(i, 2, f'{rule_name.get(k, "")}　{v} 条').font = F_BODY

    # 按附注表类型统计
    try:
        import table_types
        dist = table_types.summarize(tables)
    except Exception:
        dist = {}
    if dist:
        row2 = row + len(by_rule) + 3
        ws.cell(row2, 1, '按附注表类型分布').font = F_HEADER
        for i, (name, n) in enumerate(dist.items(), row2 + 1):
            ws.cell(i, 1, name).font = F_BODY
            ws.cell(i, 2, f'{n} 张').font = F_BODY


# ---------------------------------------------------------------- 报表勾稽

def _copy_statement_sheets(wb, report):
    """把报表的三张表复制进成果表，这样跨表公式能直接引用。"""
    import openpyxl
    src = openpyxl.load_workbook(report.path, data_only=True)
    for sheet in report.sheets:
        if sheet.name not in src.sheetnames:
            continue
        s_ws = src[sheet.name]
        name = sheet.name[:31]
        if name in wb.sheetnames:
            continue
        d_ws = wb.create_sheet(name)
        for r in range(1, s_ws.max_row + 1):
            for c in range(1, s_ws.max_column + 1):
                v = s_ws.cell(r, c).value
                if v is None:
                    continue
                cell = d_ws.cell(r, c, v)
                if isinstance(v, float):
                    cell.number_format = FMT_AMOUNT
        d_ws.column_dimensions['A'].width = 40


def _write_cross_sheet(ws, cross, stats):
    headers = ['序号', '范围', '子项目', '期间', '附注金额', '取数口径', '报表科目',
               '报表数', '差额', '状态', '说明']
    for i, h in enumerate(headers, 1):
        c = ws.cell(1, i, h)
        c.font = F_TITLE
        c.fill = FILL_TITLE
        c.alignment = Alignment(horizontal='center')

    r = 2
    for n, cc in enumerate(cross, 1):
        period = str(cc.period) if cc.period else (cc.period_key or '')
        vals = [n, cc.scope, cc.sub_item, period, cc.note_value, cc.basis,
                cc.sheet or '（未找到）', cc.report_value, None, '', cc.note]
        for i, v in enumerate(vals, 1):
            cell = ws.cell(r, i, v)
            cell.font = F_BODY
            cell.border = BORDER
            if i in (5, 8, 9):
                cell.number_format = FMT_AMOUNT

        # 差额写成活公式：附注 − 报表
        if cc.sheet and cc.rows and cc.cols:
            ws.cell(r, 9, f"=E{r}-'{cc.sheet}'!{get_column_letter(cc.cols)}{cc.rows}")
            if cc.status == 'pass':
                ws.cell(r, 9).fill = FILL_PASS
                ws.cell(r, 10, '平').font = F_BODY
                stats['cross_pass'] += 1
            elif cc.status == 'fail':
                for c in range(1, 9):
                    ws.cell(r, c).fill = FILL_FAIL_ROW
                ws.cell(r, 9).fill = FILL_FAIL_TARGET
                ws.cell(r, 10, '有差异').font = Font(name='微软雅黑', size=9, color='9C0006')
                stats['cross_fail'] += 1
            else:
                ws.cell(r, 9).fill = FILL_UNKNOWN
                ws.cell(r, 10, '待确认').font = F_BODY
                stats['cross_unmatched'] += 1
        else:
            ws.cell(r, 10, '未匹配').font = F_BODY
            ws.cell(r, 10).fill = FILL_UNKNOWN
            stats['cross_unmatched'] += 1
        r += 1

    widths = [6, 8, 20, 12, 16, 20, 30, 16, 14, 10, 26]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'


# ---------------------------------------------------------------- TB 现金流核对

def _copy_tb_sheets(wb, tb_xlsx, sheet_names):
    """把用到的 TB 现金流量表复制进成果表，让差额公式可以直接引用。"""
    import openpyxl
    if not tb_xlsx:
        return
    src = openpyxl.load_workbook(tb_xlsx, data_only=True)
    for name in sheet_names:
        if not name or name in wb.sheetnames or name not in src.sheetnames:
            continue
        s_ws = src[name]
        d_ws = wb.create_sheet(name[:31])
        for r in range(1, min(s_ws.max_row, 300) + 1):
            for c in range(1, min(s_ws.max_column, 40) + 1):
                v = s_ws.cell(r, c).value
                if v is None:
                    continue
                cell = d_ws.cell(r, c, v)
                if isinstance(v, float):
                    cell.number_format = FMT_AMOUNT
        d_ws.column_dimensions['A'].width = 46


def _write_tb_sheet(ws, tb_checks, stats):
    headers = ['序号', '章', '子项目', '项目', '期间', '附注数',
               'TB工作表', 'TB数', '差额', '状态', '说明']
    for i, h in enumerate(headers, 1):
        c = ws.cell(1, i, h)
        c.font = F_TITLE
        c.fill = FILL_TITLE
        c.alignment = Alignment(horizontal='center')

    r = 2
    for n, chk in enumerate(tb_checks, 1):
        period = str(chk.period) if chk.period else ''
        vals = [n, chk.chapter, chk.sub_item, chk.label, period, chk.note_value,
                chk.tb_sheet or '（未找到）', chk.tb_value, None, '', chk.note]
        for i, v in enumerate(vals, 1):
            cell = ws.cell(r, i, v)
            cell.font = F_BODY
            cell.border = BORDER
            if i in (6, 8, 9):
                cell.number_format = FMT_AMOUNT

        if chk.tb_sheet and chk.tb_rows and chk.tb_cols:
            ws.cell(r, 9, f"=F{r}-'{chk.tb_sheet}'!{get_column_letter(chk.tb_cols)}{chk.tb_rows}")
            if chk.status == 'pass':
                ws.cell(r, 9).fill = FILL_PASS
                ws.cell(r, 10, '平').font = F_BODY
                stats['tb_pass'] += 1
            elif chk.status == 'fail':
                for c in range(1, 9):
                    ws.cell(r, c).fill = FILL_FAIL_ROW
                ws.cell(r, 9).fill = FILL_FAIL_TARGET
                ws.cell(r, 10, '有差异').font = Font(name='微软雅黑', size=9, color='9C0006')
                stats['tb_fail'] += 1
            else:
                ws.cell(r, 9).fill = FILL_UNKNOWN
                ws.cell(r, 10, '待确认').font = F_BODY
                stats['tb_unmatched'] += 1
        else:
            ws.cell(r, 10, '未匹配').font = F_BODY
            ws.cell(r, 10).fill = FILL_UNKNOWN
            stats['tb_unmatched'] += 1
        r += 1

    widths = [6, 22, 20, 44, 12, 16, 16, 16, 14, 10, 26]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'


# ---------------------------------------------------------------- 附注号勾稽

def _write_note_no_sheet(ws, checks, stats):
    headers = ['序号', '工作表', '行号', '科目', '报表附注号', '附注实际编号', '状态', '说明']
    for i, h in enumerate(headers, 1):
        c = ws.cell(1, i, h)
        c.font = F_TITLE
        c.fill = FILL_TITLE
        c.alignment = Alignment(horizontal='center')

    r = 2
    for n, chk in enumerate(checks, 1):
        vals = [n, chk.sheet, chk.row, chk.item_name, chk.actual or '（空）',
                chk.expected, '', chk.note]
        for i, v in enumerate(vals, 1):
            cell = ws.cell(r, i, v)
            cell.font = F_BODY
            cell.border = BORDER

        st = ws.cell(r, 7)
        if chk.status == 'pass':
            ws.cell(r, 7, '一致')
            for c in (5, 6, 7):
                ws.cell(r, c).fill = FILL_PASS
            stats['nn_pass'] += 1
        elif chk.status == 'fail':
            ws.cell(r, 7, '不一致')
            for c in range(1, 8):
                ws.cell(r, c).fill = FILL_FAIL_ROW
            ws.cell(r, 5).fill = FILL_FAIL_TARGET
            ws.cell(r, 7).font = Font(name='微软雅黑', size=9, color='9C0006')
            stats['nn_fail'] += 1
        else:
            ws.cell(r, 7, '待确认')
            ws.cell(r, 7).fill = FILL_UNKNOWN
            stats['nn_unmatched'] += 1
        st.value = ws.cell(r, 7).value
        r += 1

    widths = [6, 16, 8, 40, 16, 16, 10, 34]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'


# ---------------------------------------------------------------- 未披露汇总

def _write_disclosure_sheet(ws, items, stats):
    """报表有金额、但附注找不到同名小节的科目清单，供人工判断。"""
    periods = sorted({d for it in items for d in it.amounts}, reverse=True)
    headers = ['序号', '报表', '行号', '报表科目'] + \
              [str(d) for d in periods] + ['建议']
    for i, h in enumerate(headers, 1):
        c = ws.cell(1, i, h)
        c.font = F_TITLE
        c.fill = FILL_TITLE
        c.alignment = Alignment(horizontal='center')

    note = ws.cell(2, 1, '以下科目报表上有金额，但附注主要项目里没有同名小节。'
                         '请判断是「漏披露」还是「本就不需要披露」。')
    note.font = Font(name='微软雅黑', size=9, color='B25E00')
    ws.merge_cells(start_row=2, start_column=1, end_row=2,
                   end_column=max(len(headers), 2))

    r = 3
    for n, it in enumerate(items, 1):
        vals = [n, it.sheet, it.row, it.item_name]
        for d in periods:
            vals.append(it.amounts.get(d))
        vals.append('待人工判断')
        for i, v in enumerate(vals, 1):
            cell = ws.cell(r, i, v)
            cell.font = F_BODY
            cell.border = BORDER
            if 5 <= i <= 4 + len(periods):
                cell.number_format = FMT_AMOUNT
        ws.cell(r, len(vals)).fill = FILL_UNKNOWN
        stats['missing_count'] += 1
        r += 1

    widths = [6, 16, 8, 40] + [16] * len(periods) + [16]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A3'


# ---------------------------------------------------------------- 附注内跨表

def _write_note_cross_sheet(ws, checks, stats):
    headers = ['序号', '类型', '期间', '子项目A', '表A', '金额A',
               '子项目B', '表B', '金额B', '差额', '状态', '说明']
    for i, h in enumerate(headers, 1):
        c = ws.cell(1, i, h)
        c.font = F_TITLE
        c.fill = FILL_TITLE
        c.alignment = Alignment(horizontal='center')

    r = 2
    for n, chk in enumerate(checks, 1):
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
            ws.cell(r, 11, '平').font = F_BODY
            ws.cell(r, 10).fill = FILL_PASS
            stats['nx_pass'] += 1
        elif chk.status == 'fail':
            ws.cell(r, 11, '有差异').font = Font(name='微软雅黑', size=9, color='9C0006')
            for c in range(1, 11):
                ws.cell(r, c).fill = FILL_FAIL_ROW
            ws.cell(r, 10).fill = FILL_FAIL_TARGET
            stats['nx_fail'] += 1
        else:
            ws.cell(r, 11, '待确认').font = F_BODY
            ws.cell(r, 11).fill = FILL_UNKNOWN
            stats['nx_unmatched'] += 1
        r += 1

    widths = [6, 10, 12, 20, 8, 16, 20, 8, 16, 14, 10, 40]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'


# ---------------------------------------------------------------- 入口

def write_workbook(out_path, model, tables, checks, results, cross=None,
                   report=None, tb_checks=None, tb_xlsx=None,
                   note_no_checks=None, disclosure=None, note_cross=None,
                   source_note='', log=print):
    """生成成果表。"""
    wb = Workbook()
    ws = wb.active
    ws.title = '勾稽底稿'

    ws.cell(1, 1, '财务报表附注勾稽底稿').font = Font(name='微软雅黑', size=14, bold=True,
                                              color='1F497D')
    ws.cell(2, 1, '绿色=已平（差额≤0.01）　深红=有差异　淡红底=该条有差异　'
                  '黄色=结构存疑/无法计算').font = Font(name='微软雅黑', size=9, color='666666')

    results_by_key = {id(r.item): r for r in results}
    checks_by_table = {}
    for item in checks:
        checks_by_table.setdefault((item.chapter, item.sub_item, item.table_idx),
                                   {}).setdefault(item.target_row, []).append(item)

    stats = Counter({'pass': 0, 'fail': 0, 'skip': 0, 'unclear_tables': 0,
                     'cross_pass': 0, 'cross_fail': 0, 'cross_unmatched': 0,
                     'tb_pass': 0, 'tb_fail': 0, 'tb_unmatched': 0,
                     'nn_pass': 0, 'nn_fail': 0, 'nn_unmatched': 0,
                     'nx_pass': 0, 'nx_fail': 0, 'nx_unmatched': 0,
                     'missing_count': 0})
    row_ptr = 4
    for tm in tables:
        key = (tm.chapter, tm.sub_item, tm.ref.idx)
        row_ptr = _write_table(ws, tm, checks_by_table.get(key, {}),
                               results_by_key, row_ptr, stats)
        row_ptr += 1        # 表间空一行

    ws.column_dimensions['A'].width = 46
    for c in range(2, 12):
        ws.column_dimensions[get_column_letter(c)].width = 16
    ws.freeze_panes = 'A4'

    _write_diff_sheet(wb.create_sheet('差异清单'), results)

    if report is not None:
        _copy_statement_sheets(wb, report)
    if cross:
        _write_cross_sheet(wb.create_sheet('报表勾稽'), cross, stats)
    if tb_checks is not None:
        used = sorted({c.tb_sheet for c in tb_checks if c.tb_sheet})
        _copy_tb_sheets(wb, tb_xlsx, used)
        _write_tb_sheet(wb.create_sheet('TB现金流核对'), tb_checks, stats)
    if note_no_checks is not None:
        _write_note_no_sheet(wb.create_sheet('附注号勾稽'), note_no_checks, stats)
    if note_cross is not None:
        _write_note_cross_sheet(wb.create_sheet('附注内跨表'), note_cross, stats)
    if disclosure is not None:
        _write_disclosure_sheet(wb.create_sheet('未披露汇总'), disclosure, stats)

    _write_summary_sheet(wb.create_sheet('勾稽总览'), stats, results, model, tables,
                         source_note=source_note)

    wb.save(out_path)
    log(f'[输出] 底稿已生成：{out_path}')
    log(f'[输出] 附注内部勾稽 {len(results)} 条：通过 {stats["pass"]}，'
        f'有差异 {stats["fail"]}，无法计算 {stats["skip"]}')
    if cross:
        log(f'[输出] 报表勾稽 {len(cross)} 条：平 {stats["cross_pass"]}，'
            f'有差异 {stats["cross_fail"]}，待确认 {stats["cross_unmatched"]}')
    if tb_checks:
        log(f'[输出] TB现金流核对 {len(tb_checks)} 条：平 {stats["tb_pass"]}，'
            f'有差异 {stats["tb_fail"]}，待确认 {stats["tb_unmatched"]}')
    if note_no_checks:
        log(f'[输出] 附注号勾稽 {len(note_no_checks)} 条：一致 {stats["nn_pass"]}，'
            f'不一致 {stats["nn_fail"]}，待确认 {stats["nn_unmatched"]}')
    if disclosure is not None:
        log(f'[输出] 未披露汇总 {len(disclosure)} 条（报表有金额、附注未找到同名小节，'
            f'需人工判断）')
    if note_cross:
        log(f'[输出] 附注内跨表 {len(note_cross)} 条：平 {stats["nx_pass"]}，'
            f'有差异 {stats["nx_fail"]}，待确认 {stats["nx_unmatched"]}')
    return out_path, dict(stats)
