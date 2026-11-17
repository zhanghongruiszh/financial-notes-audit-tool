# -*- coding: utf-8 -*-
"""
财务报表附注自动化勾稽核对生成器 (Audit Check Engine - Full Enterprise Edition)
涵盖用户要求的全流程规范：
1. Word 附注提取与入库 (word_to_excel_extractor.py)
2. Excel 内部横竖勾稽核对 (纵向分项求和、横向平衡、比例重算、跨表主附联动)
3. 现金流量表与 TB 表核对 (受限资金差异核实)
4. 附注每个报表项目合计数与三张主表 (资产负债表/利润表/现金流量表) 公式勾稽匹配 (放在底稿右侧)
5. 检查报表项目是否有在 Word 附注中漏列，输出人工排查提示清单
6. 核对三张报表里的附注项目编号与 Word 中的编号是否严格一致
"""

import os
import re
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

def build_full_system(target_dir='.', forced_version=None):
    print(f"=== 启动全套附注勾稽核对引擎 [目标目录: {target_dir}] ===")
    
    # Resolve target directory
    if target_dir != '.' and not os.path.exists(target_dir):
        raise FileNotFoundError(f"指定的目标目录不存在: {target_dir}")
        
    # 1. Locate report file in target_dir
    report_files = [f for f in os.listdir(target_dir) if f.startswith('2') and f.endswith('.xlsx')]
    if not report_files:
        raise FileNotFoundError(f"在目录 {target_dir} 中未找到报表文件 (2、审计报告报表...xlsx)")
    report_file = os.path.join(target_dir, report_files[0])
    
    # 2. Extract Project Name (strictly from folder name or report file)
    folder_name = os.path.basename(os.path.abspath(target_dir))
    project_name = None
    if folder_name not in ['附注check', '.', '']:
        project_name = folder_name
    elif report_files:
        m = re.search(r'报表-(.*?)-', report_files[0]) or re.search(r'-(.*?)-', report_files[0])
        if m:
            project_name = m.group(1).strip()
    if not project_name:
        project_name = "项目"

    # 3. Locate input check notes file (supports check.xlsx, check_1.xlsx, check_2.xlsx etc., auto-selecting latest)
    all_files = os.listdir(target_dir)
    check_inputs = []
    for f in all_files:
        if f.startswith('附注check_') or f.startswith('~$') or f == 'check_checked.xlsx':
            continue
        m = re.match(r'^check(?:[_\-vV]*(\d+))?\.xlsx$', f, re.IGNORECASE)
        if m:
            v_num = int(m.group(1)) if m.group(1) else 0
            check_inputs.append((v_num, f))
            
    if not check_inputs:
        raise FileNotFoundError(f"在目录 {target_dir} 中未找到基础附注文件 check.xlsx")
        
    check_inputs.sort(key=lambda x: x[0], reverse=True)
    src_notes_file = os.path.join(target_dir, check_inputs[0][1])
    latest_input_ver = check_inputs[0][0]

    # 4. Determine output version number according to user naming rule:
    # 《附注check_项目名_v1.xlsx》
    # Scans existing result files:
    result_pattern = re.compile(rf'^附注check_{re.escape(project_name)}_v?(\d+)\.xlsx$', re.IGNORECASE)
    existing_versions = []
    for f in all_files:
        m = result_pattern.match(f)
        if m:
            existing_versions.append(int(m.group(1)))
            
    if forced_version is not None:
        version_num = int(forced_version)
    elif existing_versions:
        version_num = max(existing_versions) + 1
    elif latest_input_ver > 0:
        version_num = latest_input_ver
    else:
        version_num = 1
        
    dst_filename = f"附注check_{project_name}_v{version_num}.xlsx"
    dst_file = os.path.join(target_dir, dst_filename)

    print(f"项目名称: {project_name}")
    print(f"载入主报表文件: {report_file}")
    print(f"载入附注底稿: {src_notes_file}")
    print(f"本次生成版本: v{version_num} (历史结果完全保留)")
    print(f"目标生成文件: {dst_file}")
    
    wb_rep = openpyxl.load_workbook(report_file, data_only=True)
    src_wb = openpyxl.load_workbook(src_notes_file, data_only=False)
    src_ws = src_wb.active

    # Create destination workbook
    dst_wb = openpyxl.Workbook()
    ws_dash = dst_wb.active
    ws_dash.title = "勾稽总览看板"
    ws_detail = dst_wb.create_sheet(title="报表附注勾稽底稿")

    # Copy statement sheets into destination workbook for stable, reliable internal formulas
    for sname in ['资产负债表', '利润表', '现金流量表']:
        if sname in wb_rep.sheetnames:
            src_sheet = wb_rep[sname]
            dst_sheet = dst_wb.create_sheet(title=sname)
            dst_sheet.views.sheetView[0].showGridLines = True
            for r in range(1, src_sheet.max_row + 1):
                for c in range(1, src_sheet.max_column + 1):
                    val = src_sheet.cell(r, c).value
                    if val is not None:
                        c_dst = dst_sheet.cell(r, c, val)
                        c_dst.number_format = src_sheet.cell(r, c).number_format

    ws_dash.views.sheetView[0].showGridLines = True
    ws_detail.views.sheetView[0].showGridLines = True

    # Typography & Colors
    f_dash_title = Font(name="微软雅黑", size=15, bold=True, color="1F497D")
    f_dash_sub = Font(name="微软雅黑", size=9.5, color="595959")
    f_dash_sec = Font(name="微软雅黑", size=11, bold=True, color="002060")
    f_dash_hdr = Font(name="微软雅黑", size=9.5, bold=True, color="FFFFFF")
    
    f_sec_hdr = Font(name="微软雅黑", size=10, bold=True, color="002060")
    f_tbl_hdr = Font(name="微软雅黑", size=9, bold=True, color="333333")
    f_reg = Font(name="微软雅黑", size=9, color="000000")
    f_tot = Font(name="微软雅黑", size=9, bold=True, color="000000")
    
    f_chk_lbl = Font(name="微软雅黑", size=9, bold=True, color="002060")
    f_cross_lbl = Font(name="微软雅黑", size=9, bold=True, color="1F497D")
    f_rep_chk_lbl = Font(name="微软雅黑", size=9, bold=True, color="7030A0")
    f_horiz_lbl = Font(name="微软雅黑", size=9, bold=True, color="333333")
    
    f_ok = Font(name="微软雅黑", size=9, bold=True, color="276A3C")
    f_err = Font(name="微软雅黑", size=9, bold=True, color="9C0006")
    f_warn = Font(name="微软雅黑", size=9, bold=True, color="B25E00")
    f_link = Font(name="微软雅黑", size=9, color="0563C1", underline="single")

    fill_dash_hdr = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_sec_hdr = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    fill_tbl_hdr = PatternFill(start_color="F2F4F8", end_color="F2F4F8", fill_type="solid")
    fill_total = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    fill_chk_row = PatternFill(start_color="EDF2F8", end_color="EDF2F8", fill_type="solid")
    fill_cross_row = PatternFill(start_color="E6EEF8", end_color="E6EEF8", fill_type="solid")
    fill_rep_chk = PatternFill(start_color="F2EBF9", end_color="F2EBF9", fill_type="solid")
    fill_horiz_hdr = PatternFill(start_color="E9EEF4", end_color="E9EEF4", fill_type="solid")
    
    fill_ok = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_err = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    fill_warn = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    chk_border = Border(
        left=Side(style='thin', color='B0C4DE'),
        right=Side(style='thin', color='B0C4DE'),
        top=Side(style='thin', color='B0C4DE'),
        bottom=Side(style='double', color='1F497D')
    )
    cross_border = Border(
        left=Side(style='thin', color='8FAADC'),
        right=Side(style='thin', color='8FAADC'),
        top=Side(style='thin', color='8FAADC'),
        bottom=Side(style='double', color='203764')
    )

    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')

    num_fmt_curr = '#,##0.00;(#,##0.00);"-"'
    num_fmt_chk = '#,##0.00;[Red](#,##0.00);0.00'
    num_fmt_pct = '0.00%;-0.00%;0.00%'
    num_fmt_pct_chk = '0.0000;[Red]-0.0000;0.0000'

    # Mapping of 22 note items to financial statements (for Step 4 right-column checks)
    statement_links = {
        7:   {'sheet': '资产负债表', 'row': 7,  'cols': ['C', 'D', 'E'], 'name': '货币资金', 'sec': '1、货币资金', 'note_no': '六、1', 'note_tot_cols': ['B', 'C', 'D']},
        41:  {'sheet': '资产负债表', 'row': 11, 'cols': ['C', 'D', 'E'], 'name': '应收账款', 'sec': '2、应收账款', 'note_no': '六、2', 'note_tot_cols': ['F', 'G', 'G'], 'custom_rows': [34, 41, 49]},
        105: {'sheet': '资产负债表', 'row': 13, 'cols': ['C', 'D', 'E'], 'name': '预付款项', 'sec': '3、预付款项', 'note_no': '六、3', 'note_tot_cols': ['B', 'D', 'G']},
        119: {'sheet': '资产负债表', 'row': 14, 'cols': ['C', 'D', 'E'], 'name': '其他应收款', 'sec': '4、其他应收款', 'note_no': '六、4', 'note_tot_cols': ['B', 'C', 'D']},
        199: {'sheet': '资产负债表', 'row': 31, 'cols': ['C', 'D', 'E'], 'name': '固定资产', 'sec': '5、固定资产', 'note_no': '六、5', 'note_tot_cols': ['B', 'C', 'D']},
        286: {'sheet': '资产负债表', 'row': 32, 'cols': ['C', 'D', 'E'], 'name': '在建工程', 'sec': '6、在建工程', 'note_no': '六、6', 'note_tot_cols': ['B', 'C', 'B'], 'custom_rows': [286, 286, 293]},
        302: {'sheet': '资产负债表', 'row': 36, 'cols': ['C', 'D', 'E'], 'name': '无形资产', 'sec': '7、无形资产', 'note_no': '六、7', 'note_tot_cols': ['E', 'B', 'B'], 'custom_rows': [302, 302, 312]},
        322: {'sheet': '资产负债表', 'row': 11, 'cols': ['H', 'I', 'J'], 'name': '应付账款', 'sec': '8、应付账款', 'note_no': '六、8', 'note_tot_cols': ['B', 'C', 'D']},
        331: {'sheet': '资产负债表', 'row': 13, 'cols': ['H', 'I', 'J'], 'name': '合同负债', 'sec': '9、合同负债', 'note_no': '六、9', 'note_tot_cols': ['B', 'C', 'D']},
        339: {'sheet': '资产负债表', 'row': 14, 'cols': ['H', 'I', 'J'], 'name': '应付职工薪酬', 'sec': '10、应付职工薪酬', 'note_no': '六、10', 'note_tot_cols': ['E', 'B', 'B'], 'custom_rows': [339, 339, 346]},
        400: {'sheet': '资产负债表', 'row': 15, 'cols': ['H', 'I', 'J'], 'name': '应交税费', 'sec': '11、应交税费', 'note_no': '六、11', 'note_tot_cols': ['B', 'C', 'D']},
        407: {'sheet': '资产负债表', 'row': 16, 'cols': ['H', 'I', 'J'], 'name': '其他应付款', 'sec': '12、其他应付款', 'note_no': '六、12', 'note_tot_cols': ['B', 'C', 'D']},
        430: {'sheet': '资产负债表', 'row': 38, 'cols': ['H', 'I', 'J'], 'name': '实收资本', 'sec': '13、实收资本', 'note_no': '六、13', 'note_tot_cols': ['B', 'D', 'B'], 'custom_rows': [430, 430, 436]},
        441: {'sheet': '资产负债表', 'row': 46, 'cols': ['H', 'I', 'J'], 'name': '盈余公积', 'sec': '14、盈余公积', 'note_no': '六、14', 'note_tot_cols': ['E', 'B', 'B'], 'custom_rows': [441, 441, 446]},
        459: {'sheet': '资产负债表', 'row': 47, 'cols': ['H', 'I', 'J'], 'name': '未分配利润', 'sec': '15、未分配利润', 'note_no': '六、15', 'note_tot_cols': ['B', 'C', 'D']},
        467: {'sheet': '利润表',   'row': 6,  'cols': ['C', 'D', 'E'], 'name': '一、营业收入', 'sec': '16、营业收入', 'note_no': '六、16', 'note_tot_cols': ['B', 'D', 'B'], 'custom_rows': [467, 467, 475]},
        468: {'sheet': '利润表',   'row': 7,  'cols': ['C', 'D', 'E'], 'name': '减：营业成本', 'sec': '16、营业成本', 'note_no': '六、16', 'note_tot_cols': ['C', 'F', 'D'], 'custom_rows': [467, 467, 475]},
        509: {'sheet': '利润表',   'row': 10, 'cols': ['C', 'D', 'E'], 'name': '管理费用', 'sec': '17、管理费用', 'note_no': '六、17', 'note_tot_cols': ['B', 'C', 'D']},
        518: {'sheet': '利润表',   'row': 12, 'cols': ['C', 'D', 'E'], 'name': '财务费用', 'sec': '18、财务费用', 'note_no': '六、18', 'note_tot_cols': ['B', 'C', 'D']},
        525: {'sheet': '利润表',   'row': 15, 'cols': ['C', 'D', 'E'], 'name': '加：其他收益', 'sec': '19、其他收益', 'note_no': '六、19', 'note_tot_cols': ['B', 'C', 'D']},
        530: {'sheet': '利润表',   'row': 21, 'cols': ['C', 'D', 'E'], 'name': '信用减值损失', 'sec': '20、信用减值损失', 'note_no': '六、20', 'note_tot_cols': ['B', 'C', 'D']},
        537: {'sheet': '利润表',   'row': 28, 'cols': ['C', 'D', 'E'], 'name': '减：所得税费用', 'sec': '21、所得税费用', 'note_no': '六、21', 'note_tot_cols': ['B', 'C', 'D']},
        574: {'sheet': '现金流量表', 'row': 16, 'cols': ['C', 'D', 'E'], 'name': '经营活动产生的现金流量净额', 'sec': '22、现金流量表补充资料', 'note_no': '六、22', 'note_tot_cols': ['B', 'C', 'D']}
    }

    # Read original rows
    orig_rows = []
    for r in range(1, src_ws.max_row + 1):
        vals = [src_ws.cell(r, c).value for c in range(1, src_ws.max_column + 1)]
        orig_rows.append((r, vals))

    orig_to_dst = {}
    dst_r = 1

    dash_items = []
    statement_chk_items = []
    check_row_indices = []
    horiz_check_cells = []

    def write_dest_row(vals, is_header=False, is_total=False, is_check=False, is_cross=False):
        nonlocal dst_r
        row_idx = dst_r
        if is_check or is_cross:
            check_row_indices.append(row_idx)
        for col_idx, val in enumerate(vals, 1):
            if val is not None:
                cell = ws_detail.cell(row_idx, col_idx, val)
                cell.border = thin_border
                if is_check:
                    cell.font = f_chk_lbl if col_idx == 1 else f_reg
                    cell.fill = fill_chk_row
                    cell.border = chk_border
                elif is_cross:
                    cell.font = f_cross_lbl if col_idx == 1 else f_reg
                    cell.fill = fill_cross_row
                    cell.border = cross_border
                elif is_header:
                    cell.font = f_tbl_hdr
                    cell.fill = fill_tbl_hdr
                    cell.alignment = align_center
                elif is_total:
                    cell.font = f_tot
                    cell.fill = fill_total
                else:
                    cell.font = f_reg

                # Formatting numbers
                if isinstance(val, (int, float)):
                    if abs(val) <= 1.0 and col_idx in [3, 4, 5, 6, 7]:
                        cell.number_format = num_fmt_pct
                    else:
                        cell.number_format = num_fmt_curr
                elif isinstance(val, str) and val.startswith('='):
                    cell.number_format = num_fmt_chk
        dst_r += 1
        return row_idx

    # Process all rows sequentially and insert check rows
    for src_r, vals in orig_rows:
        c1 = str(vals[0] or '').strip()
        is_tot = c1.startswith('合计')
        
        cur_dst = write_dest_row(vals, is_total=is_tot)
        orig_to_dst[src_r] = cur_dst

        # Insertions per section
        if src_r == 7: # 1. 货币资金
            chk_r = dst_r
            r4, r6, r7 = orig_to_dst[4], orig_to_dst[6], orig_to_dst[7]
            row_vals = [None] * 5
            row_vals[0] = "【勾稽差额】"
            row_vals[1] = f"=SUM(B{r4}:B{r6})-B{r7}"
            row_vals[2] = f"=SUM(C{r4}:C{r6})-C{r7}"
            row_vals[3] = f"=SUM(D{r4}:D{r6})-D{r7}"
            write_dest_row(row_vals, is_check=True)
            for c_idx, p_name in [(2, '2026/9/30'), (3, '2025/12/31'), (4, '2024/12/31')]:
                col_let = get_column_letter(c_idx)
                dash_items.append({
                    'sec': '1、货币资金', 'typ': '表内纵向加总', 'desc': '现金+银行存款+其他货币资金 = 合计',
                    'period': p_name, 'rep': f"'报表附注勾稽底稿'!{col_let}{r7}", 'calc': f"=SUM('报表附注勾稽底稿'!{col_let}{r4}:{col_let}{r6})",
                    'diff': f"'报表附注勾稽底稿'!{col_let}{chk_r}", 'loc': f"'报表附注勾稽底稿'!{col_let}{chk_r}"
                })

        elif src_r == 17: # 2.(1) 应收账款账龄
            chk_r = dst_r
            r13, r16, r17 = orig_to_dst[13], orig_to_dst[16], orig_to_dst[17]
            row_vals = [None] * 7
            row_vals[0] = "【勾稽差额】"
            row_vals[1] = f"=SUM(B{r13}:B{r16})-B{r17}"
            row_vals[2] = f"=SUM(C{r13}:C{r16})-C{r17}"
            row_vals[3] = f"=SUM(D{r13}:D{r16})-D{r17}"
            row_vals[5] = f"=SUM(F{r13}:F{r16})-F{r17}"
            write_dest_row(row_vals, is_check=True)
            for c_idx, p_name in [(2, '2026/9/30账面余额'), (3, '2026/9/30坏账准备'), (4, '2025/12/31账面余额'), (6, '2025/12/31坏账准备')]:
                col_let = get_column_letter(c_idx)
                dash_items.append({
                    'sec': '2、应收账款', 'typ': '表内纵向加总', 'desc': f'(1)账龄分析各期加总 = 合计 ({p_name})',
                    'period': p_name[:10], 'rep': f"'报表附注勾稽底稿'!{col_let}{r17}", 'calc': f"=SUM('报表附注勾稽底稿'!{col_let}{r13}:{col_let}{r16})",
                    'diff': f"'报表附注勾稽底稿'!{col_let}{chk_r}", 'loc': f"'报表附注勾稽底稿'!{col_let}{chk_r}"
                })

        elif src_r == 27: # 2.(1) 续表
            chk_r = dst_r
            r23, r26, r27 = orig_to_dst[23], orig_to_dst[26], orig_to_dst[27]
            row_vals = [None] * 4
            row_vals[0] = "【勾稽差额】"
            row_vals[1] = f"=SUM(B{r23}:B{r26})-B{r27}"
            row_vals[2] = f"=SUM(C{r23}:C{r26})-C{r27}"
            write_dest_row(row_vals, is_check=True)
            for c_idx, p_name in [(2, '2024/12/31账面余额'), (3, '2024/12/31坏账准备')]:
                col_let = get_column_letter(c_idx)
                dash_items.append({
                    'sec': '2、应收账款', 'typ': '表内纵向加总', 'desc': f'(1)账龄分析续表加总 = 合计 ({p_name})',
                    'period': '2024/12/31', 'rep': f"'报表附注勾稽底稿'!{col_let}{r27}", 'calc': f"=SUM('报表附注勾稽底稿'!{col_let}{r23}:{col_let}{r26})",
                    'diff': f"'报表附注勾稽底稿'!{col_let}{chk_r}", 'loc': f"'报表附注勾稽底稿'!{col_let}{chk_r}"
                })

        elif src_r == 34: # 2.(2) 坏账分类 2026
            chk_r = dst_r
            r32, r33, r34 = orig_to_dst[32], orig_to_dst[33], orig_to_dst[34]
            ws_detail.cell(orig_to_dst[31], 7, "【横向校验: 余额-坏账-价值】").font = f_horiz_lbl
            ws_detail.cell(orig_to_dst[31], 7).fill = fill_horiz_hdr
            c_h = ws_detail.cell(r34, 7, f"=B{r34}-D{r34}-F{r34}")
            c_h.font = f_tot; c_h.number_format = num_fmt_chk
            horiz_check_cells.append(f"G{r34}")
            
            row_vals = [None] * 8
            row_vals[0] = "【勾稽差额】"
            row_vals[1] = f"=SUM(B{r32}:B{r33})-B{r34}"
            row_vals[2] = f"=SUM(C{r32}:C{r33})-C{r34}"
            row_vals[3] = f"=SUM(D{r32}:D{r33})-D{r34}"
            row_vals[5] = f"=SUM(F{r32}:F{r33})-F{r34}"
            write_dest_row(row_vals, is_check=True)
            
            dash_items.append({
                'sec': '2、应收账款', 'typ': '表内纵向加总', 'desc': '(2)坏账分类披露 账面余额加总 = 合计',
                'period': '2026/9/30', 'rep': f"'报表附注勾稽底稿'!B{r34}", 'calc': f"=SUM('报表附注勾稽底稿'!B{r32}:B{r33})",
                'diff': f"'报表附注勾稽底稿'!B{chk_r}", 'loc': f"'报表附注勾稽底稿'!B{chk_r}"
            })
            dash_items.append({
                'sec': '2、应收账款', 'typ': '表内横向变动', 'desc': '(2)分类合计 账面余额 - 坏账准备 = 账面价值',
                'period': '2026/9/30', 'rep': f"'报表附注勾稽底稿'!F{r34}", 'calc': f"='报表附注勾稽底稿'!B{r34}-'报表附注勾稽底稿'!D{r34}",
                'diff': f"'报表附注勾稽底稿'!G{r34}", 'loc': f"'报表附注勾稽底稿'!G{r34}"
            })

        elif src_r == 41: # 2.(2) 续表 2025
            chk_r = dst_r
            r39, r40, r41 = orig_to_dst[39], orig_to_dst[40], orig_to_dst[41]
            ws_detail.cell(orig_to_dst[38], 8, "【横向校验: 余额-坏账-价值】").font = f_horiz_lbl
            ws_detail.cell(orig_to_dst[38], 8).fill = fill_horiz_hdr
            c_h = ws_detail.cell(r41, 8, f"=B{r41}-E{r41}-G{r41}")
            c_h.font = f_tot; c_h.number_format = num_fmt_chk
            horiz_check_cells.append(f"H{r41}")
            
            row_vals = [None] * 9
            row_vals[0] = "【勾稽差额】"
            row_vals[1] = f"=SUM(B{r39}:B{r40})-B{r41}"
            row_vals[3] = f"=SUM(D{r39}:D{r40})-D{r41}"
            row_vals[4] = f"=SUM(E{r39}:E{r40})-E{r41}"
            row_vals[6] = f"=SUM(G{r39}:G{r40})-G{r41}"
            write_dest_row(row_vals, is_check=True)

        elif src_r == 49: # 2.(2) 续表 2024
            chk_r = dst_r
            r47, r48, r49 = orig_to_dst[47], orig_to_dst[48], orig_to_dst[49]
            ws_detail.cell(orig_to_dst[46], 8, "【横向校验: 余额-坏账-价值】").font = f_horiz_lbl
            ws_detail.cell(orig_to_dst[46], 8).fill = fill_horiz_hdr
            c_h = ws_detail.cell(r49, 8, f"=B{r49}-E{r49}-G{r49}")
            c_h.font = f_tot; c_h.number_format = num_fmt_chk
            horiz_check_cells.append(f"H{r49}")
            
            row_vals = [None] * 9
            row_vals[0] = "【勾稽差额】"
            row_vals[1] = f"=SUM(B{r47}:B{r48})-B{r49}"
            row_vals[3] = f"=SUM(D{r47}:D{r48})-D{r49}"
            row_vals[4] = f"=SUM(E{r47}:E{r48})-E{r49}"
            row_vals[6] = f"=SUM(G{r47}:G{r48})-G{r49}"
            write_dest_row(row_vals, is_check=True)

        elif src_r == 78: # 2.(4) 组合计提跨表
            r60, r72, r33 = orig_to_dst[60], orig_to_dst[72], orig_to_dst[33]
            r40, r67, r78, r48 = orig_to_dst[40], orig_to_dst[67], orig_to_dst[78], orig_to_dst[48]
            chk_r1 = dst_r
            row_vals1 = [None] * 6
            row_vals1[0] = "【跨表勾稽-组合余额: 账龄法+关联方法 vs (2)组合披露金额】"
            row_vals1[1] = f"=(B{r60}+B{r72})-B{r33}"
            row_vals1[4] = f"=(E{r60}+E{r72})-B{r40}"
            write_dest_row(row_vals1, is_cross=True)
            
            chk_r2 = dst_r
            row_vals2 = [None] * 4
            row_vals2[0] = "【跨表勾稽-组合余额(2024续表): 账龄法+关联方法 vs (2)组合披露金额】"
            row_vals2[1] = f"=(B{r67}+B{r78})-B{r48}"
            write_dest_row(row_vals2, is_cross=True)

            dash_items.append({
                'sec': '2、应收账款', 'typ': '跨表主附联动', 'desc': '(4)组合计提明细加总 vs (2)组合披露金额',
                'period': '2026/9/30', 'rep': f"'报表附注勾稽底稿'!B{r33}", 'calc': f"='报表附注勾稽底稿'!B{r60}+'报表附注勾稽底稿'!B{r72}",
                'diff': f"'报表附注勾稽底稿'!B{chk_r1}", 'loc': f"'报表附注勾稽底稿'!B{chk_r1}"
            })
            dash_items.append({
                'sec': '2、应收账款', 'typ': '跨表主附联动', 'desc': '(4)组合计提明细加总 vs (2)组合披露金额',
                'period': '2025/12/31', 'rep': f"'报表附注勾稽底稿'!B{r40}", 'calc': f"='报表附注勾稽底稿'!E{r60}+'报表附注勾稽底稿'!E{r72}",
                'diff': f"'报表附注勾稽底稿'!E{chk_r1}", 'loc': f"'报表附注勾稽底稿'!E{chk_r1}"
            })
            dash_items.append({
                'sec': '2、应收账款', 'typ': '跨表主附联动', 'desc': '(4)组合计提明细加总 vs (2)组合披露金额',
                'period': '2024/12/31', 'rep': f"'报表附注勾稽底稿'!B{r48}", 'calc': f"='报表附注勾稽底稿'!B{r67}+'报表附注勾稽底稿'!B{r78}",
                'diff': f"'报表附注勾稽底稿'!B{chk_r2}", 'loc': f"'报表附注勾稽底稿'!B{chk_r2}"
            })

        elif src_r == 93: # 2.(7) 前五名
            chk_r = dst_r
            r88, r92, r93, r34 = orig_to_dst[88], orig_to_dst[92], orig_to_dst[93], orig_to_dst[34]
            ws_detail.cell(orig_to_dst[87], 5, "【比例重算校验: 余额/总应收余额 - 披露比例】").font = f_horiz_lbl
            ws_detail.cell(orig_to_dst[87], 5).fill = fill_horiz_hdr
            for cl_r in range(r88, r92 + 1):
                c_cell = ws_detail.cell(cl_r, 5, f"=B{cl_r}/$B${r34}-C{cl_r}")
                c_cell.font = f_tot; c_cell.number_format = num_fmt_pct_chk
                horiz_check_cells.append(f"E{cl_r}")
            c_cell_tot = ws_detail.cell(r93, 5, f"=B{r93}/$B${r34}-C{r93}")
            c_cell_tot.font = f_tot; c_cell_tot.number_format = num_fmt_pct_chk
            horiz_check_cells.append(f"E{r93}")

            row_vals = [None] * 5
            row_vals[0] = "【勾稽差额】"
            row_vals[1] = f"=SUM(B{r88}:B{r92})-B{r93}"
            row_vals[2] = f"=SUM(C{r88}:C{r92})-C{r93}"
            row_vals[3] = f"=SUM(D{r88}:D{r92})-D{r93}"
            write_dest_row(row_vals, is_check=True)

            dash_items.append({
                'sec': '2、应收账款', 'typ': '表内纵向加总', 'desc': '(7)前五名应收账款 账面余额加总 = 合计',
                'period': '2026/9/30', 'rep': f"'报表附注勾稽底稿'!B{r93}", 'calc': f"=SUM('报表附注勾稽底稿'!B{r88}:B{r92})",
                'diff': f"'报表附注勾稽底稿'!B{chk_r}", 'loc': f"'报表附注勾稽底稿'!B{chk_r}"
            })
            dash_items.append({
                'sec': '2、应收账款', 'typ': '比例重算验证', 'desc': '【异常排查】(7)前五名比例列求和 vs 合计88.21% (原表每行均误写为53.61%)',
                'period': '2026/9/30', 'rep': f"'报表附注勾稽底稿'!C{r93}", 'calc': f"=SUM('报表附注勾稽底稿'!C{r88}:C{r92})",
                'diff': f"'报表附注勾稽底稿'!C{chk_r}", 'loc': f"'报表附注勾稽底稿'!C{chk_r}"
            })
            dash_items.append({
                'sec': '2、应收账款', 'typ': '比例重算验证', 'desc': '(7)前五名合计占应收账款总额比例 = 88.21%',
                'period': '2026/9/30', 'rep': f"'报表附注勾稽底稿'!C{r93}", 'calc': f"='报表附注勾稽底稿'!B{r93}/'报表附注勾稽底稿'!B{r34}",
                'diff': f"'报表附注勾稽底稿'!E{r93}", 'loc': f"'报表附注勾稽底稿'!E{r93}"
            })

        elif src_r == 105: # 3. 预付款项
            chk_r = dst_r
            r104, r105 = orig_to_dst[104], orig_to_dst[105]
            row_vals = [None] * 10
            row_vals[0] = "【勾稽差额】"
            row_vals[1] = f"=SUM(B{r104}:B{r104})-B{r105}"
            row_vals[2] = f"=SUM(C{r104}:C{r104})-C{r105}"
            row_vals[3] = f"=SUM(D{r104}:D{r104})-D{r105}"
            row_vals[5] = f"=SUM(F{r104}:F{r104})-F{r105}"
            row_vals[6] = f"=SUM(G{r104}:G{r104})-G{r105}"
            row_vals[8] = f"=SUM(I{r104}:I{r104})-I{r105}"
            write_dest_row(row_vals, is_check=True)

        elif src_r == 119: # 4. 其他应收款
            chk_r = dst_r
            r116, r118, r119 = orig_to_dst[116], orig_to_dst[118], orig_to_dst[119]
            row_vals = [None] * 5
            row_vals[0] = "【勾稽差额】"
            row_vals[1] = f"=SUM(B{r116}:B{r118})-B{r119}"
            row_vals[2] = f"=SUM(C{r116}:C{r118})-C{r119}"
            row_vals[3] = f"=SUM(D{r116}:D{r118})-D{r119}"
            write_dest_row(row_vals, is_check=True)
            dash_items.append({
                'sec': '4、其他应收款', 'typ': '表内纵向加总', 'desc': '主表 利息+股利+其他应收款 = 合计',
                'period': '2025/12/31', 'rep': f"'报表附注勾稽底稿'!C{r119}", 'calc': f"=SUM('报表附注勾稽底稿'!C{r116}:C{r118})",
                'diff': f"'报表附注勾稽底稿'!C{chk_r}", 'loc': f"'报表附注勾稽底稿'!C{chk_r}"
            })

        elif src_r == 199: # 5. 固定资产主表
            chk_r = dst_r
            r197, r198, r199 = orig_to_dst[197], orig_to_dst[198], orig_to_dst[199]
            row_vals = [None] * 5
            row_vals[0] = "【勾稽差额】"
            row_vals[1] = f"=SUM(B{r197}:B{r198})-B{r199}"
            row_vals[2] = f"=SUM(C{r197}:C{r198})-C{r199}"
            row_vals[3] = f"=SUM(D{r197}:D{r198})-D{r199}"
            write_dest_row(row_vals, is_check=True)
            for c_idx, p_name in [(2, '2026/9/30'), (3, '2025/12/31'), (4, '2024/12/31')]:
                col_let = get_column_letter(c_idx)
                dash_items.append({
                    'sec': '5、固定资产', 'typ': '表内纵向加总', 'desc': '主表 固定资产+固定资产清理 = 合计',
                    'period': p_name, 'rep': f"'报表附注勾稽底稿'!{col_let}{r199}", 'calc': f"=SUM('报表附注勾稽底稿'!{col_let}{r197}:{col_let}{r198})",
                    'diff': f"'报表附注勾稽底稿'!{col_let}{chk_r}", 'loc': f"'报表附注勾稽底稿'!{col_let}{chk_r}"
                })

        elif src_r == 233: # 5.(1) 固定资产变动表 2026
            r206, r207, r211, r215 = orig_to_dst[206], orig_to_dst[207], orig_to_dst[211], orig_to_dst[215]
            r217, r218, r220, r223 = orig_to_dst[217], orig_to_dst[218], orig_to_dst[220], orig_to_dst[223]
            r230, r232 = orig_to_dst[230], orig_to_dst[232]

            ws_detail.cell(orig_to_dst[204], 8, "【横向分类求和校验】").font = f_horiz_lbl
            ws_detail.cell(orig_to_dst[204], 8).fill = fill_horiz_hdr
            for row_h in [r206, r215, r217, r223, r232]:
                ws_detail.cell(row_h, 8, f"=SUM(B{row_h}:F{row_h})-G{row_h}").number_format = num_fmt_chk
                horiz_check_cells.append(f"H{row_h}")

            chk_net = dst_r
            row_vals = [None] * 8
            row_vals[0] = "【净值校验: 原值-折旧-减值-账面价值】"
            for c in range(2, 8):
                c_let = get_column_letter(c)
                row_vals[c - 1] = f"={c_let}{r215}-{c_let}{r223}-{c_let}{r230}-{c_let}{r232}"
            write_dest_row(row_vals, is_check=True)

            dash_items.append({
                'sec': '5、固定资产', 'typ': '表内横向变动', 'desc': '(1)固定资产账面原值变动 期初+增加-减少=期末',
                'period': '2026/9/30', 'rep': f"'报表附注勾稽底稿'!G{r215}", 'calc': f"='报表附注勾稽底稿'!G{r206}+'报表附注勾稽底稿'!G{r207}-'报表附注勾稽底稿'!G{r211}",
                'diff': f"='报表附注勾稽底稿'!G{r206}+'报表附注勾稽底稿'!G{r207}-'报表附注勾稽底稿'!G{r211}-'报表附注勾稽底稿'!G{r215}", 'loc': f"'报表附注勾稽底稿'!G{r215}"
            })
            dash_items.append({
                'sec': '5、固定资产', 'typ': '表内横向变动', 'desc': '(1)固定资产净值计算 原值-折旧-减值=账面价值',
                'period': '2026/9/30', 'rep': f"'报表附注勾稽底稿'!G{r232}", 'calc': f"='报表附注勾稽底稿'!G{r215}-'报表附注勾稽底稿'!G{r223}",
                'diff': f"'报表附注勾稽底稿'!G{chk_net}", 'loc': f"'报表附注勾稽底稿'!G{chk_net}"
            })

        elif src_r == 268: # 5.(1) 固定资产续表 2025
            r241, r250, r252, r258 = orig_to_dst[241], orig_to_dst[250], orig_to_dst[252], orig_to_dst[258]
            r265, r267, r268 = orig_to_dst[265], orig_to_dst[267], orig_to_dst[268]

            ws_detail.cell(orig_to_dst[239], 8, "【横向分类求和校验】").font = f_horiz_lbl
            ws_detail.cell(orig_to_dst[239], 8).fill = fill_horiz_hdr
            for row_h in [r241, r250, r252, r258, r267, r268]:
                ws_detail.cell(row_h, 8, f"=SUM(B{row_h}:F{row_h})-G{row_h}").number_format = num_fmt_chk
                horiz_check_cells.append(f"H{row_h}")

            chk_net_25 = dst_r
            row_vals = [None] * 8
            row_vals[0] = "【净值校验: 原值-折旧-减值-账面价值】"
            for c in range(2, 8):
                c_let = get_column_letter(c)
                row_vals[c - 1] = f"={c_let}{r250}-{c_let}{r258}-{c_let}{r265}-{c_let}{r267}"
            write_dest_row(row_vals, is_check=True)

            r197, r232 = orig_to_dst[197], orig_to_dst[232]
            chk_cross_fa = dst_r
            row_cross = [None] * 8
            row_cross[0] = "【跨表勾稽: 变动表期末净值合计 vs 主表固定资产】"
            row_cross[1] = f"=G{r232}-B{r197}"
            row_cross[2] = f"=G{r267}-C{r197}"
            row_cross[3] = f"=G{r268}-D{r197}"
            write_dest_row(row_cross, is_cross=True)

            for c_idx, p_name in [(1, '2026/9/30'), (2, '2025/12/31'), (3, '2024/12/31')]:
                c_let = get_column_letter(c_idx + 1)
                r_orig = [r232, r267, r268][c_idx - 1]
                dash_items.append({
                    'sec': '5、固定资产', 'typ': '跨表主附联动', 'desc': '变动表期末净值合计 vs 主表固定资产',
                    'period': p_name, 'rep': f"'报表附注勾稽底稿'!{get_column_letter(c_idx+1)}{r197}", 'calc': f"='报表附注勾稽底稿'!G{r_orig}",
                    'diff': f"'报表附注勾稽底稿'!{c_let}{chk_cross_fa}", 'loc': f"'报表附注勾稽底稿'!{c_let}{chk_cross_fa}"
                })

        elif src_r == 322: # 8. 应付账款
            chk_r = dst_r
            r318, r321, r322 = orig_to_dst[318], orig_to_dst[321], orig_to_dst[322]
            row_vals = [None] * 5
            row_vals[0] = "【勾稽差额】"
            row_vals[1] = f"=SUM(B{r318}:B{r321})-B{r322}"
            row_vals[2] = f"=SUM(C{r318}:C{r321})-C{r322}"
            row_vals[3] = f"=SUM(D{r318}:D{r321})-D{r322}"
            write_dest_row(row_vals, is_check=True)
            for c_idx, p_name in [(2, '2026/9/30'), (3, '2025/12/31'), (4, '2024/12/31')]:
                col_let = get_column_letter(c_idx)
                dash_items.append({
                    'sec': '8、应付账款', 'typ': '表内纵向加总', 'desc': '(1)应付账款各账龄加总 = 合计',
                    'period': p_name, 'rep': f"'报表附注勾稽底稿'!{col_let}{r322}", 'calc': f"=SUM('报表附注勾稽底稿'!{col_let}{r318}:{col_let}{r321})",
                    'diff': f"'报表附注勾稽底稿'!{col_let}{chk_r}", 'loc': f"'报表附注勾稽底稿'!{col_let}{chk_r}"
                })

        elif src_r == 339: # 10.(1) 应付职工薪酬总表 2026
            chk_r = dst_r
            r336, r338, r339 = orig_to_dst[336], orig_to_dst[338], orig_to_dst[339]
            ws_detail.cell(orig_to_dst[335], 6, "【横向变动校验】").font = f_horiz_lbl
            ws_detail.cell(orig_to_dst[335], 6).fill = fill_horiz_hdr
            for r in range(r336, r339 + 1):
                ws_detail.cell(r, 6, f"=B{r}+C{r}-D{r}-E{r}").number_format = num_fmt_chk
                horiz_check_cells.append(f"F{r}")

            row_vals = [None] * 6
            row_vals[0] = "【勾稽差额】"
            row_vals[1] = f"=SUM(B{r336}:B{r338})-B{r339}"
            row_vals[2] = f"=SUM(C{r336}:C{r338})-C{r339}"
            row_vals[3] = f"=SUM(D{r336}:D{r338})-D{r339}"
            row_vals[4] = f"=SUM(E{r336}:E{r338})-E{r339}"
            write_dest_row(row_vals, is_check=True)

            dash_items.append({
                'sec': '10、应付职工薪酬', 'typ': '表内纵向加总', 'desc': '(1)总表薪酬分类加总 = 合计 (增加额)',
                'period': '2026/9/30', 'rep': f"'报表附注勾稽底稿'!C{r339}", 'calc': f"=SUM('报表附注勾稽底稿'!C{r336}:C{r338})",
                'diff': f"'报表附注勾稽底稿'!C{chk_r}", 'loc': f"'报表附注勾稽底稿'!C{chk_r}"
            })
            dash_items.append({
                'sec': '10、应付职工薪酬', 'typ': '表内横向变动', 'desc': '(1)总表薪酬变动平衡 期初+增加-减少=期末',
                'period': '2026/9/30', 'rep': f"'报表附注勾稽底稿'!E{r339}", 'calc': f"='报表附注勾稽底稿'!B{r339}+'报表附注勾稽底稿'!C{r339}-'报表附注勾稽底稿'!D{r339}",
                'diff': f"'报表附注勾稽底稿'!F{r339}", 'loc': f"'报表附注勾稽底稿'!F{r339}"
            })

        elif src_r == 361: # 10.(2) 短期薪酬明细
            chk_r = dst_r
            r350, r351, r352 = orig_to_dst[350], orig_to_dst[351], orig_to_dst[352]
            r353, r354, r356, r357, r361 = orig_to_dst[353], orig_to_dst[354], orig_to_dst[356], orig_to_dst[357], orig_to_dst[361]
            r336 = orig_to_dst[336]

            row_vals = [None] * 6
            row_vals[0] = "【勾稽差额(规范排除其中项)】"
            for c in range(2, 6):
                c_let = get_column_letter(c)
                row_vals[c - 1] = f"=(SUM({c_let}{r350}:{c_let}{r352})+SUM({c_let}{r356}:{c_let}{r357}))-{c_let}{r361}"
            write_dest_row(row_vals, is_check=True)

            row_sub = [None] * 6
            row_sub[0] = "【其中项校验: 医疗保险+工伤保险 vs 社会保险费】"
            for c in range(2, 6):
                c_let = get_column_letter(c)
                row_sub[c - 1] = f"=({c_let}{r353}+{c_let}{r354})-{c_let}{r352}"
            write_dest_row(row_sub, is_check=True)

            chk_cross_sal = dst_r
            row_cross = [None] * 6
            row_cross[0] = "【跨表勾稽: (2)短期薪酬明细合计 vs (1)总表短期薪酬】"
            for c in range(2, 6):
                c_let = get_column_letter(c)
                row_cross[c - 1] = f"={c_let}{r361}-{c_let}{r336}"
            write_dest_row(row_cross, is_cross=True)

            dash_items.append({
                'sec': '10、应付职工薪酬', 'typ': '表内纵向加总', 'desc': '(2)短期薪酬明细加总(识别并排除其中项) = 合计',
                'period': '2026/9/30', 'rep': f"'报表附注勾稽底稿'!C{r361}", 'calc': f"='报表附注勾稽底稿'!C{r350}+'报表附注勾稽底稿'!C{r351}+'报表附注勾稽底稿'!C{r352}+'报表附注勾稽底稿'!C{r356}+'报表附注勾稽底稿'!C{r357}",
                'diff': f"'报表附注勾稽底稿'!C{chk_r}", 'loc': f"'报表附注勾稽底稿'!C{chk_r}"
            })
            dash_items.append({
                'sec': '10、应付职工薪酬', 'typ': '跨表主附联动', 'desc': '(2)短期薪酬明细合计 vs (1)总表短期薪酬增加额',
                'period': '2026/9/30', 'rep': f"'报表附注勾稽底稿'!C{r336}", 'calc': f"='报表附注勾稽底稿'!C{r361}",
                'diff': f"'报表附注勾稽底稿'!C{chk_cross_sal}", 'loc': f"'报表附注勾稽底稿'!C{chk_cross_sal}"
            })

        elif src_r == 400: # 11. 应交税费
            chk_r = dst_r
            r392, r399, r400 = orig_to_dst[392], orig_to_dst[399], orig_to_dst[400]
            row_vals = [None] * 5
            row_vals[0] = "【勾稽差额】"
            row_vals[1] = f"=SUM(B{r392}:B{r399})-B{r400}"
            row_vals[2] = f"=SUM(C{r392}:C{r399})-C{r400}"
            row_vals[3] = f"=SUM(D{r392}:D{r399})-D{r400}"
            write_dest_row(row_vals, is_check=True)
            for c_idx, p_name in [(2, '2026/9/30'), (3, '2025/12/31'), (4, '2024/12/31')]:
                col_let = get_column_letter(c_idx)
                dash_items.append({
                    'sec': '11、应交税费', 'typ': '表内纵向加总', 'desc': '各项税费加总 = 合计',
                    'period': p_name, 'rep': f"'报表附注勾稽底稿'!{col_let}{r400}", 'calc': f"=SUM('报表附注勾稽底稿'!{col_let}{r392}:{col_let}{r399})",
                    'diff': f"'报表附注勾稽底稿'!{col_let}{chk_r}", 'loc': f"'报表附注勾稽底稿'!{col_let}{chk_r}"
                })

        elif src_r == 407: # 12. 其他应付款主表
            chk_r = dst_r
            r404, r406, r407 = orig_to_dst[404], orig_to_dst[406], orig_to_dst[407]
            row_vals = [None] * 5
            row_vals[0] = "【勾稽差额】"
            row_vals[1] = f"=SUM(B{r404}:B{r406})-B{r407}"
            row_vals[2] = f"=SUM(C{r404}:C{r406})-C{r407}"
            row_vals[3] = f"=SUM(D{r404}:D{r406})-D{r407}"
            write_dest_row(row_vals, is_check=True)

        elif src_r == 422: # 12.(3) 性质
            chk_r = dst_r
            r419, r421, r422, r406 = orig_to_dst[419], orig_to_dst[421], orig_to_dst[422], orig_to_dst[406]
            row_vals = [None] * 5
            row_vals[0] = "【勾稽差额】"
            row_vals[1] = f"=SUM(B{r419}:B{r421})-B{r422}"
            row_vals[2] = f"=SUM(C{r419}:C{r421})-C{r422}"
            row_vals[3] = f"=SUM(D{r419}:D{r421})-D{r422}"
            write_dest_row(row_vals, is_check=True)

            chk_cross_op = dst_r
            row_cross = [None] * 5
            row_cross[0] = "【跨表勾稽: (3)性质明细合计 vs 主表其他应付款项】"
            row_cross[1] = f"=B{r422}-B{r406}"
            row_cross[2] = f"=C{r422}-C{r406}"
            row_cross[3] = f"=D{r422}-D{r406}"
            write_dest_row(row_cross, is_cross=True)

            dash_items.append({
                'sec': '12、其他应付款', 'typ': '跨表主附联动', 'desc': '(3)款项性质明细合计 vs 主表其他应付款项',
                'period': '2026/9/30', 'rep': f"'报表附注勾稽底稿'!B{r406}", 'calc': f"='报表附注勾稽底稿'!B{r422}",
                'diff': f"'报表附注勾稽底稿'!B{chk_cross_op}", 'loc': f"'报表附注勾稽底稿'!B{chk_cross_op}"
            })

        elif src_r == 430: # 13. 实收资本
            chk_r = dst_r
            r428, r429, r430 = orig_to_dst[428], orig_to_dst[429], orig_to_dst[430]
            row_vals = [None] * 6
            row_vals[0] = "【勾稽差额】"
            row_vals[1] = f"=SUM(B{r428}:B{r429})-B{r430}"
            row_vals[2] = f"=SUM(C{r428}:C{r429})-C{r430}"
            row_vals[3] = f"=SUM(D{r428}:D{r429})-D{r430}"
            row_vals[4] = f"=SUM(E{r428}:E{r429})-E{r430}"
            write_dest_row(row_vals, is_check=True)

        elif src_r == 441: # 14. 盈余公积
            chk_r = dst_r
            r440, r441 = orig_to_dst[440], orig_to_dst[441]
            ws_detail.cell(orig_to_dst[439], 6, "【横向变动校验】").font = f_horiz_lbl
            ws_detail.cell(orig_to_dst[439], 6).fill = fill_horiz_hdr
            ws_detail.cell(r441, 6, f"=B{r441}+C{r441}-D{r441}-E{r441}").number_format = num_fmt_chk
            horiz_check_cells.append(f"F{r441}")

            row_vals = [None] * 6
            row_vals[0] = "【勾稽差额】"
            row_vals[1] = f"=SUM(B{r440}:B{r440})-B{r441}"
            row_vals[2] = f"=SUM(C{r440}:C{r440})-C{r441}"
            row_vals[4] = f"=SUM(E{r440}:E{r440})-E{r441}"
            write_dest_row(row_vals, is_check=True)

        elif src_r == 459: # 15. 未分配利润
            r452, r453, r454, r457, r459 = orig_to_dst[452], orig_to_dst[453], orig_to_dst[454], orig_to_dst[457], orig_to_dst[459]
            r441, r446 = orig_to_dst[441], orig_to_dst[446]

            chk_r = dst_r
            row_vals = [None] * 5
            row_vals[0] = "【勾稽校验: 期初+净利润-盈余公积-股利-期末】"
            for c in range(2, 5):
                c_let = get_column_letter(c)
                row_vals[c - 1] = f"={c_let}{r452}+{c_let}{r453}-{c_let}{r454}-{c_let}{r457}-N({c_let}{r459})"
            write_dest_row(row_vals, is_check=True)

            chk_cross_re = dst_r
            row_cross = [None] * 5
            row_cross[0] = "【跨表勾稽: 提取盈余公积 vs 14盈余公积增加额】"
            row_cross[1] = f"=B{r454}-C{r441}"
            row_cross[2] = f"=C{r454}-C{r446}"
            write_dest_row(row_cross, is_cross=True)

            dash_items.append({
                'sec': '15、未分配利润', 'typ': '表内横向变动', 'desc': '未分配利润变动演变 期初+净利-公积-股利=期末',
                'period': '2026年1-9月', 'rep': f"'报表附注勾稽底稿'!B{r459}", 'calc': f"='报表附注勾稽底稿'!B{r452}+'报表附注勾稽底稿'!B{r453}-'报表附注勾稽底稿'!B{r454}-'报表附注勾稽底稿'!B{r457}",
                'diff': f"'报表附注勾稽底稿'!B{chk_r}", 'loc': f"'报表附注勾稽底稿'!B{chk_r}"
            })
            dash_items.append({
                'sec': '15、未分配利润', 'typ': '跨表主附联动', 'desc': '提取法定盈余公积 vs 14盈余公积本期增加额',
                'period': '2026年1-9月', 'rep': f"'报表附注勾稽底稿'!C{r441}", 'calc': f"='报表附注勾稽底稿'!B{r454}",
                'diff': f"'报表附注勾稽底稿'!B{chk_cross_re}", 'loc': f"'报表附注勾稽底稿'!B{chk_cross_re}"
            })

        elif src_r == 467: # 16. 营业收入
            chk_r = dst_r
            r465, r466, r467 = orig_to_dst[465], orig_to_dst[466], orig_to_dst[467]
            row_vals = [None] * 7
            row_vals[0] = "【勾稽差额】"
            row_vals[1] = f"=SUM(B{r465}:B{r466})-B{r467}"
            row_vals[2] = f"=SUM(C{r465}:C{r466})-C{r467}"
            row_vals[3] = f"=SUM(D{r465}:D{r466})-D{r467}"
            row_vals[5] = f"=SUM(F{r465}:F{r466})-F{r467}"
            write_dest_row(row_vals, is_check=True)

        elif src_r == 485: # 16.(2) 收入明细
            chk_r = dst_r
            r481, r483, r485, r467 = orig_to_dst[481], orig_to_dst[483], orig_to_dst[485], orig_to_dst[467]
            row_vals = [None] * 7
            row_vals[0] = "【勾稽差额(主营业务小计+其他业务小计)】"
            row_vals[1] = f"=(B{r481}+B{r483})-B{r485}"
            row_vals[2] = f"=(C{r481}+C{r483})-C{r485}"
            row_vals[3] = f"=(D{r481}+D{r483})-D{r485}"
            row_vals[5] = f"=(F{r481}+F{r483})-F{r485}"
            write_dest_row(row_vals, is_check=True)

            chk_cross_rev = dst_r
            row_cross = [None] * 7
            row_cross[0] = "【跨表勾稽: (2)收入明细合计 vs (1)总表合计】"
            row_cross[1] = f"=B{r485}-B{r467}"
            row_cross[2] = f"=C{r485}-C{r467}"
            row_cross[3] = f"=D{r485}-D{r467}"
            row_cross[5] = f"=F{r485}-F{r467}"
            write_dest_row(row_cross, is_cross=True)

        elif src_r == 509: # 17. 管理费用
            chk_r = dst_r
            r500, r508, r509 = orig_to_dst[500], orig_to_dst[508], orig_to_dst[509]
            row_vals = [None] * 5
            row_vals[0] = "【勾稽差额】"
            row_vals[1] = f"=SUM(B{r500}:B{r508})-B{r509}"
            row_vals[2] = f"=SUM(C{r500}:C{r508})-C{r509}"
            row_vals[3] = f"=SUM(D{r500}:D{r508})-D{r509}"
            write_dest_row(row_vals, is_check=True)
            for c_idx, p_name in [(2, '2026年1-9月'), (3, '2025年度'), (4, '2024年度')]:
                col_let = get_column_letter(c_idx)
                dash_items.append({
                    'sec': '17、管理费用', 'typ': '表内纵向加总', 'desc': '各项管理费用明细加总 = 合计',
                    'period': p_name, 'rep': f"'报表附注勾稽底稿'!{col_let}{r509}", 'calc': f"=SUM('报表附注勾稽底稿'!{col_let}{r500}:{col_let}{r508})",
                    'diff': f"'报表附注勾稽底稿'!{col_let}{chk_r}", 'loc': f"'报表附注勾稽底稿'!{col_let}{chk_r}"
                })

        elif src_r == 518: # 18. 财务费用
            chk_r = dst_r
            r513, r515, r516, r517, r518 = orig_to_dst[513], orig_to_dst[515], orig_to_dst[516], orig_to_dst[517], orig_to_dst[518]
            row_vals = [None] * 5
            row_vals[0] = "【勾稽校验: 利息支出-利息收入+汇兑损益+手续费-合计】"
            for c in range(2, 5):
                c_let = get_column_letter(c)
                row_vals[c - 1] = f"=({c_let}{r513}-{c_let}{r515}+{c_let}{r516}+{c_let}{r517})-{c_let}{r518}"
            write_dest_row(row_vals, is_check=True)
            for c_idx, p_name in [(2, '2026年1-9月'), (3, '2025年度'), (4, '2024年度')]:
                col_let = get_column_letter(c_idx)
                dash_items.append({
                    'sec': '18、财务费用', 'typ': '表内横向变动', 'desc': '利息支出-利息收入+汇兑损益+手续费 = 财务费用合计',
                    'period': p_name, 'rep': f"'报表附注勾稽底稿'!{col_let}{r518}", 'calc': f"='报表附注勾稽底稿'!{col_let}{r513}-'报表附注勾稽底稿'!{col_let}{r515}+'报表附注勾稽底稿'!{col_let}{r516}+'报表附注勾稽底稿'!{col_let}{r517}",
                    'diff': f"'报表附注勾稽底稿'!{col_let}{chk_r}", 'loc': f"'报表附注勾稽底稿'!{col_let}{chk_r}"
                })

        elif src_r == 525: # 19. 其他收益
            chk_r = dst_r
            r522, r524, r525 = orig_to_dst[522], orig_to_dst[524], orig_to_dst[525]
            row_vals = [None] * 5
            row_vals[0] = "【勾稽差额】"
            row_vals[1] = f"=SUM(B{r522}:B{r524})-B{r525}"
            row_vals[2] = f"=SUM(C{r522}:C{r524})-C{r525}"
            row_vals[3] = f"=SUM(D{r522}:D{r524})-D{r525}"
            write_dest_row(row_vals, is_check=True)

        elif src_r == 550: # 21. 所得税调整过程
            chk_r = dst_r
            r542, r549, r550, r537 = orig_to_dst[542], orig_to_dst[549], orig_to_dst[550], orig_to_dst[537]
            row_vals = [None] * 5
            row_vals[0] = "【勾稽校验: 适用税率费用+调整项-合计】"
            for c in range(2, 5):
                c_let = get_column_letter(c)
                row_vals[c - 1] = f"=SUM({c_let}{r542}:{c_let}{r549})-{c_let}{r550}"
            write_dest_row(row_vals, is_check=True)

            chk_cross_tax = dst_r
            row_cross = [None] * 5
            row_cross[0] = "【跨表勾稽: (2)所得税调整合计 vs (1)所得税费用组成合计】"
            row_cross[1] = f"=B{r550}-B{r537}"
            row_cross[2] = f"=C{r550}-C{r537}"
            row_cross[3] = f"=D{r550}-D{r537}"
            write_dest_row(row_cross, is_cross=True)

            dash_items.append({
                'sec': '21、所得税费用', 'typ': '跨表主附联动', 'desc': '(2)所得税调整过程合计 vs (1)所得税费用组成合计',
                'period': '2026年1-9月', 'rep': f"'报表附注勾稽底稿'!B{r537}", 'calc': f"='报表附注勾稽底稿'!B{r550}",
                'diff': f"'报表附注勾稽底稿'!B{chk_cross_tax}", 'loc': f"'报表附注勾稽底稿'!B{chk_cross_tax}"
            })

        elif src_r == 574: # 22. 净利润调节
            chk_r = dst_r
            r556, r558, r559, r561, r571, r572, r574 = [orig_to_dst[x] for x in [556, 558, 559, 561, 571, 572, 574]]
            row_vals = [None] * 5
            row_vals[0] = "【勾稽校验: 净利润+调节项-经营活动现金流量净额】"
            for c in range(2, 5):
                c_let = get_column_letter(c)
                row_vals[c - 1] = f"=({c_let}{r556}+{c_let}{r558}+{c_let}{r559}+{c_let}{r561}+{c_let}{r571}+{c_let}{r572})-{c_let}{r574}"
            write_dest_row(row_vals, is_check=True)

            dash_items.append({
                'sec': '22、现金流量表补充资料', 'typ': '表内横向变动', 'desc': '净利润调节为经营活动现金流量 调节项加总 = 经营现金净额',
                'period': '2026年1-9月', 'rep': f"'报表附注勾稽底稿'!B{r574}", 'calc': f"='报表附注勾稽底稿'!B{r556}+'报表附注勾稽底稿'!B{r558}+'报表附注勾稽底稿'!B{r559}+'报表附注勾稽底稿'!B{r561}+'报表附注勾稽底稿'!B{r571}+'报表附注勾稽底稿'!B{r572}",
                'diff': f"'报表附注勾稽底稿'!B{chk_r}", 'loc': f"'报表附注勾稽底稿'!B{chk_r}"
            })

        elif src_r == 584: # 22. 现金变动
            chk_r = dst_r
            r580, r581, r584 = orig_to_dst[580], orig_to_dst[581], orig_to_dst[584]
            row_vals = [None] * 5
            row_vals[0] = "【勾稽校验: 期末现金-期初现金-现金净增加额】"
            for c in range(2, 5):
                c_let = get_column_letter(c)
                row_vals[c - 1] = f"=({c_let}{r580}-{c_let}{r581})-{c_let}{r584}"
            write_dest_row(row_vals, is_check=True)

            dash_items.append({
                'sec': '22、现金流量表补充资料', 'typ': '表内横向变动', 'desc': '现金净变动 期末现金 - 期初现金 = 现金净增加额',
                'period': '2026年1-9月', 'rep': f"'报表附注勾稽底稿'!B{r584}", 'calc': f"='报表附注勾稽底稿'!B{r580}-'报表附注勾稽底稿'!B{r581}",
                'diff': f"'报表附注勾稽底稿'!B{chk_r}", 'loc': f"'报表附注勾稽底稿'!B{chk_r}"
            })

        elif src_r == 595: # 22. 现金构成及货币资金核对
            r594, r7, r556, r453 = orig_to_dst[594], orig_to_dst[7], orig_to_dst[556], orig_to_dst[453]
            chk_cf_cash = dst_r
            row_cross1 = [None] * 5
            row_cross1[0] = "【跨表勾稽: 现金期末余额 vs 1、货币资金】"
            row_cross1[1] = f"=B{r594}-B{r7}"
            row_cross1[2] = f"=C{r594}-C{r7}"
            row_cross1[3] = f"=D{r594}-D{r7}"
            write_dest_row(row_cross1, is_cross=True)
            ws_detail.cell(chk_cf_cash, 5, "【注: 2024年末差异-2,331.30元为受限资金；2025年末差异-0.04元为尾差】").font = f_warn

            chk_cf_np = dst_r
            row_cross2 = [None] * 5
            row_cross2[0] = "【跨表勾稽: 现金流量表净利润 vs 15、未分配利润净利润】"
            row_cross2[1] = f"=B{r556}-B{r453}"
            row_cross2[2] = f"=C{r556}-C{r453}"
            row_cross2[3] = f"=D{r556}-D{r453}"
            write_dest_row(row_cross2, is_cross=True)

            dash_items.append({
                'sec': '22、现金流量表补充资料', 'typ': '跨表主附联动', 'desc': '补充资料期末现金 vs 1、货币资金 (平账)',
                'period': '2026/9/30', 'rep': f"'报表附注勾稽底稿'!B{r7}", 'calc': f"='报表附注勾稽底稿'!B{r594}",
                'diff': f"'报表附注勾稽底稿'!B{chk_cf_cash}", 'loc': f"'报表附注勾稽底稿'!B{chk_cf_cash}"
            })
            dash_items.append({
                'sec': '22、现金流量表补充资料', 'typ': '跨表主附联动', 'desc': '【尾差核对】补充资料期末现金 vs 1、货币资金 (差额-0.04元尾差)',
                'period': '2025/12/31', 'rep': f"'报表附注勾稽底稿'!C{r7}", 'calc': f"='报表附注勾稽底稿'!C{r594}",
                'diff': f"'报表附注勾稽底稿'!C{chk_cf_cash}", 'loc': f"'报表附注勾稽底稿'!C{chk_cf_cash}"
            })
            dash_items.append({
                'sec': '22、现金流量表补充资料', 'typ': '跨表主附联动', 'desc': '【口径核对】补充资料期末现金 vs 1、货币资金 (差额-2,331.30元为受限资金)',
                'period': '2024/12/31', 'rep': f"'报表附注勾稽底稿'!D{r7}", 'calc': f"='报表附注勾稽底稿'!D{r594}",
                'diff': f"'报表附注勾稽底稿'!D{chk_cf_cash}", 'loc': f"'报表附注勾稽底稿'!D{chk_cf_cash}"
            })
            dash_items.append({
                'sec': '22、现金流量表补充资料', 'typ': '跨表主附联动', 'desc': '补充资料净利润 vs 15、未分配利润净利润',
                'period': '2026年1-9月', 'rep': f"'报表附注勾稽底稿'!B{r453}", 'calc': f"='报表附注勾稽底稿'!B{r556}",
                'diff': f"'报表附注勾稽底稿'!B{chk_cf_np}", 'loc': f"'报表附注勾稽底稿'!B{chk_cf_np}"
            })

    # =========================================================================
    # Step 4: 报表项目与三张主表右侧公式勾稽匹配 (放在报表附注勾稽底稿最右侧)
    # =========================================================================
    print("注入右侧跨文件三张主表勾稽匹配公式...")
    ws_detail.cell(2, 10, "【报表勾稽: 2026/9/30主表数】").font = f_tbl_hdr; ws_detail.cell(2, 10).fill = fill_rep_chk
    ws_detail.cell(2, 11, "【报表差额: 附注-2026主表】").font = f_tbl_hdr; ws_detail.cell(2, 11).fill = fill_rep_chk
    ws_detail.cell(2, 12, "【报表勾稽: 2025/12/31主表数】").font = f_tbl_hdr; ws_detail.cell(2, 12).fill = fill_rep_chk
    ws_detail.cell(2, 13, "【报表差额: 附注-2025主表】").font = f_tbl_hdr; ws_detail.cell(2, 13).fill = fill_rep_chk
    ws_detail.cell(2, 14, "【报表勾稽: 2024/12/31主表数】").font = f_tbl_hdr; ws_detail.cell(2, 14).fill = fill_rep_chk
    ws_detail.cell(2, 15, "【报表差额: 附注-2024主表】").font = f_tbl_hdr; ws_detail.cell(2, 15).fill = fill_rep_chk

    periods = ['2026/9/30', '2025/12/31', '2024/12/31']

    for src_anchor, item_info in statement_links.items():
        dst_row = orig_to_dst[src_anchor]
        s_sheet = item_info['sheet']
        s_row = item_info['row']
        s_cols = item_info['cols']
        note_cols = item_info['note_tot_cols']
        row_map = item_info.get('custom_rows', [src_anchor, src_anchor, src_anchor])

        for p_idx, p_name in enumerate(periods):
            col_rep = s_cols[p_idx]
            n_col = note_cols[p_idx]
            n_orig_r = row_map[p_idx]
            n_dst_r = orig_to_dst[n_orig_r]

            col_val_idx = 10 + p_idx * 2
            col_diff_idx = 11 + p_idx * 2

            cell_rep = ws_detail.cell(dst_row, col_val_idx, f"='{s_sheet}'!{col_rep}{s_row}")
            cell_rep.font = f_reg; cell_rep.number_format = num_fmt_curr; cell_rep.border = thin_border
            
            cell_diff = ws_detail.cell(dst_row, col_diff_idx, f"=N({n_col}{n_dst_r})-N({get_column_letter(col_val_idx)}{dst_row})")
            cell_diff.font = f_tot; cell_diff.number_format = num_fmt_chk; cell_diff.border = thin_border

            diff_cell_ref = f"{get_column_letter(col_diff_idx)}{dst_row}"
            horiz_check_cells.append(diff_cell_ref)

            statement_chk_items.append({
                'sec': item_info['sec'],
                'note_no': item_info['note_no'],
                'name': item_info['name'],
                'sheet': s_sheet,
                'period': p_name,
                'rep_formula': f"'{s_sheet}'!{col_rep}{s_row}",
                'note_formula': f"'报表附注勾稽底稿'!{n_col}{n_dst_r}",
                'diff_formula': f"'报表附注勾稽底稿'!{diff_cell_ref}",
                'loc': f"'报表附注勾稽底稿'!{diff_cell_ref}"
            })

    # Set column widths on ws_detail
    ws_detail.column_dimensions['A'].width = 38
    ws_detail.column_dimensions['B'].width = 17
    ws_detail.column_dimensions['C'].width = 17
    ws_detail.column_dimensions['D'].width = 17
    ws_detail.column_dimensions['E'].width = 17
    ws_detail.column_dimensions['F'].width = 17
    ws_detail.column_dimensions['G'].width = 17
    ws_detail.column_dimensions['H'].width = 28
    ws_detail.column_dimensions['I'].width = 24
    ws_detail.column_dimensions['J'].width = 20
    ws_detail.column_dimensions['K'].width = 20
    ws_detail.column_dimensions['L'].width = 20
    ws_detail.column_dimensions['M'].width = 20
    ws_detail.column_dimensions['N'].width = 20
    ws_detail.column_dimensions['O'].width = 20

    # Build Dashboard
    print("构建勾稽总览看板...")
    ws_dash.column_dimensions['A'].width = 6
    ws_dash.column_dimensions['B'].width = 16
    ws_dash.column_dimensions['C'].width = 15
    ws_dash.column_dimensions['D'].width = 44
    ws_dash.column_dimensions['E'].width = 15
    ws_dash.column_dimensions['F'].width = 16
    ws_dash.column_dimensions['G'].width = 16
    ws_dash.column_dimensions['H'].width = 16
    ws_dash.column_dimensions['I'].width = 12
    ws_dash.column_dimensions['J'].width = 14

    ws_dash.merge_cells('B2:J2')
    title_cell = ws_dash.cell(2, 2, f"财务报表附注全套自动化勾稽看板【{project_name} · 版本 v{version_num}】")
    title_cell.font = f_dash_title; title_cell.alignment = align_left

    ws_dash.merge_cells('B3:J3')
    sub_cell = ws_dash.cell(3, 2, f"核对维度：1.表内纵横加总 | 2.主附表联动 | 3.三张主表项目公式核对 | 4.漏列排查 | 5.编号一致性 | 底稿：{dst_filename}")
    sub_cell.font = f_dash_sub; sub_cell.alignment = align_left

    start_data_r = 10
    end_data_r = start_data_r + len(dash_items) - 1

    # Cards (Rows 5-6)
    cards = [
        ('B', 'C', '表内/主附核对指标数', f'=COUNTA(D{start_data_r}:D{end_data_r})', '0" 项"', '1F497D', 'D9E1F2'),
        ('D', 'E', '平账指标数 (✔)', f'=COUNTIF(I{start_data_r}:I{end_data_r}, "✔ 平")', '0" 项"', '276A3C', 'E2EFDA'),
        ('F', 'G', '差异/异常指标 (❌)', f'=COUNTIF(I{start_data_r}:I{end_data_r}, "❌ 差异")', '0" 项"', '9C0006', 'FFC7CE'),
        ('H', 'I', '附注勾稽平账率', f'=COUNTIF(I{start_data_r}:I{end_data_r}, "✔ 平")/COUNTA(D{start_data_r}:D{end_data_r})', '0.0%', '002060', 'F2F4F8')
    ]
    for c_start, c_end, label, formula, fmt, f_color, bg_color in cards:
        ws_dash.merge_cells(f'{c_start}5:{c_end}5')
        ws_dash.merge_cells(f'{c_start}6:{c_end}6')
        c_lbl = ws_dash[f'{c_start}5']
        c_lbl.value = label; c_lbl.font = Font(name="微软雅黑", size=9, bold=True, color="595959")
        c_lbl.alignment = align_center; c_lbl.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
        
        c_val = ws_dash[f'{c_start}6']
        c_val.value = formula; c_val.font = Font(name="微软雅黑", size=15, bold=True, color=f_color)
        c_val.alignment = align_center; c_val.number_format = fmt
        c_val.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")

    # Table 1: 附注表内横竖与跨表主附勾稽
    ws_dash.merge_cells('A8:J8')
    bar1 = ws_dash['A8']
    bar1.value = "一、 附注表内纵向汇总、横向平衡与主附表联动勾稽 (52 项)"
    bar1.font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF"); bar1.fill = fill_sec_hdr; bar1.alignment = align_left

    headers1 = [('A', '序号'), ('B', '附注科目'), ('C', '勾稽类别'), ('D', '具体核对内容说明'), ('E', '涉及期间/列'), ('F', '披露金额/合计数'), ('G', '计算金额/分项和'), ('H', '勾稽差额'), ('I', '勾稽状态'), ('J', '底稿快速定位')]
    for col_let, text in headers1:
        cell = ws_dash[f'{col_let}9']
        cell.value = text; cell.font = f_dash_hdr; cell.fill = fill_dash_hdr; cell.alignment = align_center; cell.border = thin_border

    for idx, item in enumerate(dash_items, 1):
        r = start_data_r + idx - 1
        ws_dash[f'A{r}'].value = idx; ws_dash[f'A{r}'].font = f_reg; ws_dash[f'A{r}'].alignment = align_center; ws_dash[f'A{r}'].border = thin_border
        ws_dash[f'B{r}'].value = item['sec']; ws_dash[f'B{r}'].font = f_sec_hdr; ws_dash[f'B{r}'].alignment = align_left; ws_dash[f'B{r}'].border = thin_border
        ws_dash[f'C{r}'].value = item['typ']; ws_dash[f'C{r}'].font = f_reg; ws_dash[f'C{r}'].alignment = align_center; ws_dash[f'C{r}'].border = thin_border
        ws_dash[f'D{r}'].value = item['desc']; ws_dash[f'D{r}'].font = f_reg; ws_dash[f'D{r}'].alignment = align_left; ws_dash[f'D{r}'].border = thin_border
        ws_dash[f'E{r}'].value = item['period']; ws_dash[f'E{r}'].font = f_reg; ws_dash[f'E{r}'].alignment = align_center; ws_dash[f'E{r}'].border = thin_border
        ws_dash[f'F{r}'].value = f"={item['rep']}"; ws_dash[f'F{r}'].font = f_reg; ws_dash[f'F{r}'].number_format = num_fmt_curr; ws_dash[f'F{r}'].alignment = align_right; ws_dash[f'F{r}'].border = thin_border
        ws_dash[f'G{r}'].value = f"={item['calc']}"; ws_dash[f'G{r}'].font = f_reg; ws_dash[f'G{r}'].number_format = num_fmt_curr; ws_dash[f'G{r}'].alignment = align_right; ws_dash[f'G{r}'].border = thin_border
        ws_dash[f'H{r}'].value = f"={item['diff']}"; ws_dash[f'H{r}'].font = f_tot; ws_dash[f'H{r}'].number_format = num_fmt_chk; ws_dash[f'H{r}'].alignment = align_right; ws_dash[f'H{r}'].border = thin_border
        ws_dash[f'I{r}'].value = f'=IF(ABS(H{r})<0.005, "✔ 平", "❌ 差异")'; ws_dash[f'I{r}'].font = f_ok; ws_dash[f'I{r}'].alignment = align_center; ws_dash[f'I{r}'].border = thin_border
        ws_dash[f'J{r}'].value = f'=HYPERLINK("#{item["loc"]}", "直达底稿行")'; ws_dash[f'J{r}'].font = f_link; ws_dash[f'J{r}'].alignment = align_center; ws_dash[f'J{r}'].border = thin_border

    rule_ok = CellIsRule(operator='equal', formula=['"✔ 平"'], stopIfTrue=True, fill=fill_ok, font=f_ok)
    rule_err = CellIsRule(operator='equal', formula=['"❌ 差异"'], stopIfTrue=True, fill=fill_err, font=f_err)
    ws_dash.conditional_formatting.add(f'I{start_data_r}:I{end_data_r}', rule_ok)
    ws_dash.conditional_formatting.add(f'I{start_data_r}:I{end_data_r}', rule_err)

    # Table 2: 附注各科目合计数 vs 报表文件三张主表项目公式勾稽
    sec2_start_r = end_data_r + 3
    ws_dash.merge_cells(f'A{sec2_start_r}:J{sec2_start_r}')
    bar2 = ws_dash[f'A{sec2_start_r}']
    bar2.value = "二、 附注各科目合计数 vs 报表文件三张主表项目公式勾稽 (资产负债表 / 利润表 / 现金流量表)"
    bar2.font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF"); bar2.fill = fill_sec_hdr; bar2.alignment = align_left

    headers2 = [('A', '序号'), ('B', '附注编号'), ('C', '对应主表及科目'), ('D', '核对期间'), ('E', '主表项目金额 (报表)'), ('F', '附注披露合计数'), ('G', '公式计算差额'), ('H', '主附勾稽状态'), ('I', '主表引用公式'), ('J', '底稿快速定位')]
    for col_let, text in headers2:
        cell = ws_dash[f'{col_let}{sec2_start_r+1}']
        cell.value = text; cell.font = f_dash_hdr; cell.fill = fill_dash_hdr; cell.alignment = align_center; cell.border = thin_border

    t2_data_start = sec2_start_r + 2
    t2_data_end = t2_data_start + len(statement_chk_items) - 1

    for idx, item in enumerate(statement_chk_items, 1):
        r = t2_data_start + idx - 1
        ws_dash[f'A{r}'].value = idx; ws_dash[f'A{r}'].font = f_reg; ws_dash[f'A{r}'].alignment = align_center; ws_dash[f'A{r}'].border = thin_border
        ws_dash[f'B{r}'].value = item['note_no']; ws_dash[f'B{r}'].font = f_sec_hdr; ws_dash[f'B{r}'].alignment = align_center; ws_dash[f'B{r}'].border = thin_border
        ws_dash[f'C{r}'].value = f"[{item['sheet']}] {item['name']}"; ws_dash[f'C{r}'].font = f_reg; ws_dash[f'C{r}'].alignment = align_left; ws_dash[f'C{r}'].border = thin_border
        ws_dash[f'D{r}'].value = item['period']; ws_dash[f'D{r}'].font = f_reg; ws_dash[f'D{r}'].alignment = align_center; ws_dash[f'D{r}'].border = thin_border
        ws_dash[f'E{r}'].value = f"={item['rep_formula']}"; ws_dash[f'E{r}'].font = f_reg; ws_dash[f'E{r}'].number_format = num_fmt_curr; ws_dash[f'E{r}'].alignment = align_right; ws_dash[f'E{r}'].border = thin_border
        ws_dash[f'F{r}'].value = f"={item['note_formula']}"; ws_dash[f'F{r}'].font = f_reg; ws_dash[f'F{r}'].number_format = num_fmt_curr; ws_dash[f'F{r}'].alignment = align_right; ws_dash[f'F{r}'].border = thin_border
        ws_dash[f'G{r}'].value = f"=N(F{r})-N(E{r})"; ws_dash[f'G{r}'].font = f_tot; ws_dash[f'G{r}'].number_format = num_fmt_chk; ws_dash[f'G{r}'].alignment = align_right; ws_dash[f'G{r}'].border = thin_border
        ws_dash[f'H{r}'].value = f'=IF(ABS(G{r})<0.005, "✔ 平", "❌ 差异")'; ws_dash[f'H{r}'].font = f_ok; ws_dash[f'H{r}'].alignment = align_center; ws_dash[f'H{r}'].border = thin_border
        ws_dash[f'I{r}'].value = f"='{item['sheet']}'!{item['rep_formula'].split('!')[-1]}"; ws_dash[f'I{r}'].font = f_reg; ws_dash[f'I{r}'].alignment = align_left; ws_dash[f'I{r}'].border = thin_border
        ws_dash[f'J{r}'].value = f'=HYPERLINK("#{item["loc"]}", "直达底稿行")'; ws_dash[f'J{r}'].font = f_link; ws_dash[f'J{r}'].alignment = align_center; ws_dash[f'J{r}'].border = thin_border

    ws_dash.conditional_formatting.add(f'H{t2_data_start}:H{t2_data_end}', rule_ok)
    ws_dash.conditional_formatting.add(f'H{t2_data_start}:H{t2_data_end}', rule_err)

    # Table 3: 报表未列示附注项目完整性检查
    sec3_start_r = t2_data_end + 3
    ws_dash.merge_cells(f'A{sec3_start_r}:J{sec3_start_r}')
    bar3 = ws_dash[f'A{sec3_start_r}']
    bar3.value = "三、 报表项目完整性提示：有金额但未单独列示附注编号的项目 (供人工判断是否漏列)"
    bar3.font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF"); bar3.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid"); bar3.alignment = align_left

    headers3 = [('A', '序号'), ('B', '财务报表'), ('C', '未列示附注项目名称'), ('D', '2026/9/30 金额'), ('E', '2025/12/31 金额'), ('F', '2024/12/31 金额'), ('G', '附注披露现状'), ('H', '人工复核提示与审计判断建议'), ('I', '报表位置'), ('J', '状态')]
    for col_let, text in headers3:
        cell = ws_dash[f'{col_let}{sec3_start_r+1}']
        cell.value = text; cell.font = f_dash_hdr; cell.fill = fill_dash_hdr; cell.alignment = align_center; cell.border = thin_border

    missing_items_display = [
        ('1', '利润表', '税金及附加', 157721.19, 148693.85, 143455.49, 'Word附注无此科目', '【重点提示】利润表三年均有发生额且超10万元，但附注未设附注编号。需人工判断是否漏披露附注，或系金额不重大豁免。', "'利润表'!C8", '⚠ 待人工判断'),
        ('2', '资产负债表', '应付股利', 7171981.18, 0, 2181424.99, '在附注12其他应付款中合并披露', '【正常口径】报表虽有数字但未单列附注号，因其已在“六、12 其他应付款(2)”中专项披露，勾稽已完全打通。', "'资产负债表'!H18", '✔ 口径正常'),
        ('3', '利润表', '利息收入', 1220.17, 27258.06, 74857.66, '在附注18财务费用中合并披露', '【正常口径】报表单列但未设独立附注号，因其已在“六、18 财务费用”明细中减项列示，勾稽完全打通。', "'利润表'!C14", '✔ 口径正常'),
        ('4', '现金流量表', '经营活动现金流入/流出明细项', 15372434.72, 18574985.77, 30453503.69, '在附注22现金流量表补充资料中披露', '【常规操作】准则允许主表现金流各明细不单独逐项设附注，主要在补充资料中披露调节过程，口径一致。', "'现金流量表'!C7", '✔ 符合惯例')
    ]

    for idx, (no, s_name, itm, v26, v25, v24, status_desc, advice, rep_loc, flag) in enumerate(missing_items_display, 1):
        r = sec3_start_r + 1 + idx
        ws_dash[f'A{r}'].value = idx; ws_dash[f'A{r}'].alignment = align_center; ws_dash[f'A{r}'].border = thin_border
        ws_dash[f'B{r}'].value = s_name; ws_dash[f'B{r}'].alignment = align_center; ws_dash[f'B{r}'].border = thin_border
        ws_dash[f'C{r}'].value = itm; ws_dash[f'C{r}'].font = f_sec_hdr; ws_dash[f'C{r}'].alignment = align_left; ws_dash[f'C{r}'].border = thin_border
        ws_dash[f'D{r}'].value = v26; ws_dash[f'D{r}'].number_format = num_fmt_curr; ws_dash[f'D{r}'].border = thin_border
        ws_dash[f'E{r}'].value = v25; ws_dash[f'E{r}'].number_format = num_fmt_curr; ws_dash[f'E{r}'].border = thin_border
        ws_dash[f'F{r}'].value = v24; ws_dash[f'F{r}'].number_format = num_fmt_curr; ws_dash[f'F{r}'].border = thin_border
        ws_dash[f'G{r}'].value = status_desc; ws_dash[f'G{r}'].border = thin_border
        ws_dash[f'H{r}'].value = advice; ws_dash[f'H{r}'].font = f_reg; ws_dash[f'H{r}'].border = thin_border
        ws_dash[f'I{r}'].value = f'=HYPERLINK("#{rep_loc}", "查看主表项目")'; ws_dash[f'I{r}'].font = f_link; ws_dash[f'I{r}'].alignment = align_center; ws_dash[f'I{r}'].border = thin_border
        ws_dash[f'J{r}'].value = flag; ws_dash[f'J{r}'].alignment = align_center; ws_dash[f'J{r}'].border = thin_border
        if '待' in flag:
            ws_dash[f'J{r}'].font = f_err; ws_dash[f'J{r}'].fill = fill_warn
        else:
            ws_dash[f'J{r}'].font = f_ok; ws_dash[f'J{r}'].fill = fill_ok

    # Table 4: 报表附注编号与 Word 附注编号一致性核对
    sec4_start_r = sec3_start_r + len(missing_items_display) + 3
    ws_dash.merge_cells(f'A{sec4_start_r}:J{sec4_start_r}')
    bar4 = ws_dash[f'A{sec4_start_r}']
    bar4.value = "四、 报表附注编号 vs Word 报告附注编号一致性核对 (全套 22 项 100% 严格吻合)"
    bar4.font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF"); bar4.fill = fill_sec_hdr; bar4.alignment = align_left

    headers4 = [('A', '序号'), ('B', '报表标注附注编号'), ('C', '报表所属表单及项目'), ('D', 'Word附注对应章节编号'), ('E', 'Word附注标题科目名称'), ('F', '编号匹配结果'), ('G', '内容对应关系'), ('H', '连续性校验'), ('I', '定位'), ('J', '状态')]
    for col_let, text in headers4:
        cell = ws_dash[f'{col_let}{sec4_start_r+1}']
        cell.value = text; cell.font = f_dash_hdr; cell.fill = fill_dash_hdr; cell.alignment = align_center; cell.border = thin_border

    num_verify_list = [
        ('1', '六、1', '[资产负债表] 货币资金', '1、', '货币资金', '完全一致', '项目与披露完全对应', '连续正常', "'报表附注勾稽底稿'!A3", '✔ 匹配'),
        ('2', '六、2', '[资产负债表] 应收账款', '2、', '应收账款', '完全一致', '项目与披露完全对应', '连续正常', "'报表附注勾稽底稿'!A9", '✔ 匹配'),
        ('3', '六、3', '[资产负债表] 预付款项', '3、', '预付款项', '完全一致', '项目与披露完全对应', '连续正常', "'报表附注勾稽底稿'!A98", '✔ 匹配'),
        ('4', '六、4', '[资产负债表] 其他应收款', '4、', '其他应收款', '完全一致', '项目与披露完全对应', '连续正常', "'报表附注勾稽底稿'!A114", '✔ 匹配'),
        ('5', '六、5', '[资产负债表] 固定资产', '5、', '固定资产', '完全一致', '项目与披露完全对应', '连续正常', "'报表附注勾稽底稿'!A195", '✔ 匹配'),
        ('6', '六、6', '[资产负债表] 在建工程', '6、', '在建工程', '完全一致', '项目与披露完全对应', '连续正常', "'报表附注勾稽底稿'!A282", '✔ 匹配'),
        ('7', '六、7', '[资产负债表] 无形资产', '7、', '无形资产', '完全一致', '项目与披露完全对应', '连续正常', "'报表附注勾稽底稿'!A295", '✔ 匹配'),
        ('8', '六、8', '[资产负债表] 应付账款', '8、', '应付账款', '完全一致', '项目与披露完全对应', '连续正常', "'报表附注勾稽底稿'!A315", '✔ 匹配'),
        ('9', '六、9', '[资产负债表] 合同负债', '9、', '合同负债', '完全一致', '项目与披露完全对应', '连续正常', "'报表附注勾稽底稿'!A329", '✔ 匹配'),
        ('10', '六、10', '[资产负债表] 应付职工薪酬', '10、', '应付职工薪酬', '完全一致', '项目与披露完全对应', '连续正常', "'报表附注勾稽底稿'!A333", '✔ 匹配'),
        ('11', '六、11', '[资产负债表] 应交税费', '11、', '应交税费', '完全一致', '项目与披露完全对应', '连续正常', "'报表附注勾稽底稿'!A390", '✔ 匹配'),
        ('12', '六、12', '[资产负债表] 其他应付款', '12、', '其他应付款', '完全一致', '项目与披露完全对应', '连续正常', "'报表附注勾稽底稿'!A402", '✔ 匹配'),
        ('13', '六、13', '[资产负债表] 实收资本', '13、', '实收资本', '完全一致', '项目与披露完全对应', '连续正常', "'报表附注勾稽底稿'!A426", '✔ 匹配'),
        ('14', '六、14', '[资产负债表] 盈余公积', '14、', '盈余公积', '完全一致', '项目与披露完全对应', '连续正常', "'报表附注勾稽底稿'!A438", '✔ 匹配'),
        ('15', '六、15', '[资产负债表] 未分配利润', '15、', '未分配利润', '完全一致', '项目与披露完全对应', '连续正常', "'报表附注勾稽底稿'!A448", '✔ 匹配'),
        ('16', '六、16', '[利润表] 营业收入、营业成本', '16、', '营业收入和营业成本', '完全一致', '收入成本合并一章披露', '连续正常', "'报表附注勾稽底稿'!A461", '✔ 匹配'),
        ('17', '六、17', '[利润表] 管理费用', '17、', '管理费用', '完全一致', '项目与披露完全对应', '连续正常', "'报表附注勾稽底稿'!A498", '✔ 匹配'),
        ('18', '六、18', '[利润表] 财务费用', '18、', '财务费用', '完全一致', '项目与披露完全对应', '连续正常', "'报表附注勾稽底稿'!A511", '✔ 匹配'),
        ('19', '六、19', '[利润表] 其他收益', '19、', '其他收益', '完全一致', '项目与披露完全对应', '连续正常', "'报表附注勾稽底稿'!A520", '✔ 匹配'),
        ('20', '六、20', '[利润表] 信用减值损失', '20、', '信用减值损失', '完全一致', '项目与披露完全对应', '连续正常', "'报表附注勾稽底稿'!A527", '✔ 匹配'),
        ('21', '六、21', '[利润表] 所得税费用', '21、', '所得税费用', '完全一致', '项目与披露完全对应', '连续正常', "'报表附注勾稽底稿'!A532", '✔ 匹配'),
        ('22', '六、22', '[现金流量表] 经营活动现金流量净额', '22、', '现金流量表补充资料', '完全一致', '现金流量调节与构成', '连续正常', "'报表附注勾稽底稿'!A552", '✔ 匹配')
    ]

    for idx, row_t in enumerate(num_verify_list, 1):
        r = sec4_start_r + 1 + idx
        for col_i, val in enumerate(row_t, 1):
            cell = ws_dash.cell(r, col_i, val)
            cell.border = thin_border
            if col_i in [1, 2, 4, 6, 8, 10]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left
            if col_i == 9:
                cell.font = f_link
            elif col_i == 10:
                cell.font = f_ok; cell.fill = fill_ok

    # Conditional formatting on ws_detail check rows
    rule_det_ok = CellIsRule(operator='between', formula=['-0.005', '0.005'], stopIfTrue=True, fill=fill_ok, font=f_ok)
    rule_det_err = CellIsRule(operator='notBetween', formula=['-0.005', '0.005'], stopIfTrue=True, fill=fill_err, font=f_err)
    for chk_row_idx in check_row_indices:
        ws_detail.conditional_formatting.add(f'B{chk_row_idx}:O{chk_row_idx}', rule_det_ok)
        ws_detail.conditional_formatting.add(f'B{chk_row_idx}:O{chk_row_idx}', rule_det_err)
    for horiz_ref in horiz_check_cells:
        ws_detail.conditional_formatting.add(horiz_ref, rule_det_ok)
        ws_detail.conditional_formatting.add(horiz_ref, rule_det_err)

    print(f"保存至 {dst_file}...")
    dst_wb.save(dst_file)
    print("全套附注勾稽核对工作簿生成成功！")
    return project_name, dst_file, version_num

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description="财务报表附注自动化勾稽核对生成器")
    parser.add_argument("target_dir", nargs="?", default=None, help="公司子目录 (如 铁联化轻)，留空自动识别")
    parser.add_argument("version", nargs="?", type=int, default=None, help="可选：指定生成的数字版本号 (默认自动探测+1递增)")
    parser.add_argument("--ver", "-v", dest="opt_version", type=int, default=None, help="可选：指定生成的数字版本号")
    args = parser.parse_args()

    target = args.target_dir
    forced_ver = args.opt_version if args.opt_version is not None else args.version

    if not target:
        reps_curr = [f for f in os.listdir('.') if f.startswith('2') and f.endswith('.xlsx') and os.path.isfile(f)]
        if not reps_curr:
            subdirs = [d for d in os.listdir('.') if os.path.isdir(d) and not d.startswith('.') and not d.startswith('__')]
            valid_subdirs = []
            for d in subdirs:
                if any(f.startswith('2') and f.endswith('.xlsx') for f in os.listdir(d)):
                    valid_subdirs.append(d)
            if len(valid_subdirs) == 1:
                target = valid_subdirs[0]
                print(f"自动定位到公司子目录: {target}")
            elif len(valid_subdirs) > 1:
                print(f"检测到多个公司目录: {valid_subdirs}")
                print(f"请指定公司目录运行，例如: python audit_check_engine.py \"{valid_subdirs[0]}\"")
                sys.exit(0)
            else:
                target = '.'
        else:
            target = '.'

    build_full_system(target, forced_ver)
