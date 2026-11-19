# -*- coding: utf-8 -*-
"""
verify_order.py — 用报表「附注」列反查 Word 编号顺序是否自洽。

用户的判断（本工具要验证的规律）：
  1. Word 附注的编号顺序 = 资产 → 负债 → 权益 → 利润 → 现流
  2. 报表里项目的排列顺序与该编号顺序一致（单调递增）
  3. 报表标注的编号 ↔ Word 实际编号，应当一一对应且名称相符

三项检查：
  A. 报表内编号是否单调递增（每张报表各查一遍）
  B. 报表编号在 Word 里是否存在同名小节（号 → 名 对得上）
  C. Word 里有编号、但报表没引用的小节（可能是漏标或多余披露）

用法：
    python tools/verify_order.py <报表.xlsx> <附注.docx> [--out docs/顺序校验.md]
"""

import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl                         # noqa: E402
import notes_parser                     # noqa: E402
import report_reader as rr              # noqa: E402


def cn_to_int(text):
    """'六、18' / '18' / '（三）' 里取出数字部分，统一成 int。"""
    if text is None:
        return None
    s = str(text)
    m = re.search(r'(\d+)\s*$', s)
    if m:
        return int(m.group(1))
    return None


def report_notes(path):
    """[(报表, 分栏列, 行号, 项目名, 附注号原文, 附注号数字)]"""
    book = rr.read_report(path)
    wb = openpyxl.load_workbook(path, data_only=True)
    out = []
    for sheet in book.sheets:
        if sheet.name not in wb.sheetnames:
            continue
        ws = wb[sheet.name]
        for name, (row, item_col) in sheet.items.items():
            # 附注号一般紧挨在项目名右侧
            note_text = None
            for dc in (1, 2, -1, 3):
                if item_col + dc < 1:
                    continue
                cell = ws.cell(row, item_col + dc).value
                if isinstance(cell, str) and re.search(r'[一二三四五六七八九十\d]+[、.]?\s*\d*', cell) \
                        and len(cell) <= 12:
                    note_text = cell.strip()
                    break
            if note_text:
                out.append((sheet.name, item_col, row, name, note_text,
                            cn_to_int(note_text)))
    return out


def word_items(path, work):
    model = notes_parser.parse_notes(path, work_dir=work)
    rows = []
    for ch in model.chapters:
        for sub in ch.sub_items:
            rows.append((ch.name, getattr(sub, 'number', None), sub.name))
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('report')
    ap.add_argument('notes')
    ap.add_argument('--out', default='docs/顺序校验.md')
    ap.add_argument('--work', default='_work/verify')
    args = ap.parse_args()

    notes_rows = report_notes(args.report)
    words = word_items(args.notes, args.work)

    # Word：编号 -> 小节名
    by_no = {}
    for ch, no, name in words:
        if no:
            by_no.setdefault(int(no), []).append(name)

    lines = ['# 顺序校验（报表附注号 ↔ Word 编号）\n']
    lines.append(f'- 报表标注编号的项目：{len(notes_rows)} 个')
    lines.append(f'- Word 带编号的小节：{len(by_no)} 个\n')

    # A. 单调性
    lines.append('## A. 报表内编号是否单调递增\n')
    # 资产负债表是左右双栏，必须**按栏**分别看顺序，
    # 否则会把右栏（负债）的编号和左栏（资产）混在一起比
    by_sheet = {}
    for sh, icol, row, name, text, num in notes_rows:
        if num is not None:
            by_sheet.setdefault((sh, icol), []).append((row, num, name))
    for (sh, icol), items in sorted(by_sheet.items()):
        items.sort()
        breaks = [(items[i][2], items[i][1], items[i + 1][2], items[i + 1][1])
                  for i in range(len(items) - 1) if items[i + 1][1] < items[i][1]]
        lines.append(f'- {sh}（第{icol}栏）：{len(items)} 项，'
                     + ('顺序正常' if not breaks else f'**{len(breaks)} 处倒序**'))
        for a, na, b, nb in breaks[:5]:
            lines.append(f'    - {a}（{na}）→ {b}（{nb}）')

    # B. 号 → 名 是否对得上
    lines.append('\n## B. 报表编号在 Word 里是否存在同名小节\n')
    bad = 0
    for sh, icol, row, name, text, num in notes_rows:
        if num is None:
            continue
        names = by_no.get(num, [])
        if not names:
            lines.append(f'- **{sh}·{name}** 标注 {text}，但 Word 里没有编号 {num} 的小节')
            bad += 1
        elif not any(rr.normalize_name(name)[:2] in rr.normalize_name(n)
                     or rr.normalize_name(n)[:2] in rr.normalize_name(name)
                     for n in names):
            lines.append(f'- {sh}·{name} 标注 {text} → Word 是「{names[0]}」，名称对不上')
            bad += 1
    lines.append(f'\n对不上的：{bad} 个')

    # D. 子项必然与主项共用附注编号
    lines.append('\n## D. 子项与主项是否共用编号\n')
    book = rr.read_report(args.report)
    wb = openpyxl.load_workbook(args.report, data_only=True)
    sub_bad = 0
    sub_ok = 0
    for sheet in book.sheets:
        if sheet.name not in wb.sheetnames:
            continue
        ws = wb[sheet.name]
        # 按栏收集：项目名/行号/缩进/编号
        col_items = {}
        for name, (row, icol) in sheet.items.items():
            raw = ws.cell(row, icol).value
            if not isinstance(raw, str):
                continue
            indent = len(raw) - len(raw.lstrip())
            # 只有"其中/减：/加："开头的才算子项；光靠缩进会误判（模板行缩进也很深）
            is_child = raw.strip().startswith(('其中', '减：', '加：'))
            # 没有任何金额的模板行不参与校验
            has_money = any(
                isinstance(ws.cell(row, icol + d).value, (int, float))
                for d in range(1, 5) if icol + d >= 1)
            if not has_money:
                continue
            note = None
            for dc in (1, 2, -1, 3):
                if icol + dc < 1:
                    continue
                cell = ws.cell(row, icol + dc).value
                if isinstance(cell, str) and len(cell) <= 12 and re.search(r'\d', cell):
                    note = cell.strip()
                    break
            col_items.setdefault(icol, []).append((row, name, indent, is_child, note))
        for icol, items in col_items.items():
            items.sort()
            last_parent = None
            for row, name, indent, is_child, note in items:
                if not is_child:
                    last_parent = (name, note)
                    continue
                if last_parent is None:
                    continue
                pname, pnote = last_parent
                if note and pnote:
                    if cn_to_int(note) == cn_to_int(pnote):
                        sub_ok += 1
                    else:
                        lines.append(f'- {sheet.name}「{name}」{note} ≠ 主项「{pname}」{pnote}')
                        sub_bad += 1
                elif not note:
                    lines.append(f'- {sheet.name}「{name}」没编号，主项「{pname}」是 {pnote}')
                    sub_bad += 1
    lines.append(f'\n一致 {sub_ok} 个，不一致/缺失 {sub_bad} 个')

    # C. Word 有、报表没引用
    used = {t[5] for t in notes_rows if t[5] is not None}
    lines.append('\n## C. Word 有编号但报表没引用的小节\n')
    extra = [f'{no}、{names[0]}' for no, names in sorted(by_no.items())
             if no not in used]
    for e in extra[:20]:
        lines.append(f'- {e}')
    lines.append(f'\n合计：{len(extra)} 个')

    os.makedirs(os.path.dirname(args.out) or '.', exist_ok=True)
    with open(args.out, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines) + '\n')
    print('\n'.join(lines[:18]))
    print(f'\n已写入 {args.out}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
