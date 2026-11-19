# -*- coding: utf-8 -*-
"""
note_number_checks.py — 报表主表「附注」列 与 附注实际编号 的一致性勾稽。

背景
----
报表（资产负债表/利润表/现金流量表）的「附注」列写着该科目对应的附注号，
如 `六、3`。这一列很容易和附注正文对不上：附注重排过、换了年份模板、
或者干脆是空的。

这里**只做核对、不改文件**（改文件用 report_numbering.fix_report_numbers）。

编号口径
--------
如果先做过「Word附注编号重排」，勾稽应基于**重排后的**附注编号——
这由调用方决定传入哪份附注（pipeline 会优先使用 `_编号已规范.docx`）。
"""

import datetime
import re
from dataclasses import dataclass

import openpyxl
from openpyxl.cell.cell import MergedCell

import notes_parser
import report_numbering as rn
import report_reader as rr

# 「六、3」「6、3」「六.3」「(六)3」「六-3」都能认
RE_NOTE_REF = re.compile(
    r'^\s*[（(]?\s*([一二三四五六七八九十百]+|\d+)\s*[)）]?\s*[、.．,，\-—_]?\s*(\d+)\s*$')


@dataclass
class NoteNoCheck:
    sheet: str
    row: int
    item_name: str
    actual: str
    expected: str
    chapter_num: object = None
    item_num: object = None
    status: str = 'unmatched'      # pass / fail / unmatched
    note: str = ''


def parse_note_ref(text):
    """把「六、3」解析成 (章号, 项目号)。解析不出返回 (None, None)。"""
    s = str(text or '').strip()
    if not s:
        return None, None
    m = RE_NOTE_REF.match(s)
    if not m:
        return None, None
    head = m.group(1)
    item = int(m.group(2))
    if head.isdigit():
        chapter = int(head)
    else:
        chapter = notes_parser.cn_to_int(head)
    return chapter, item


def format_ref(chapter, number):
    """生成「六、3」这样的规范写法。章号里已带的顿号不会重复。"""
    prefix = (chapter.number_text or '').strip()
    if not prefix and chapter.number is not None:
        prefix = str(chapter.number)
    if not prefix:
        return str(number)
    if prefix[-1] in '、.．,，-—':
        return f'{prefix}{number}'
    return f'{prefix}、{number}'


def build_note_no_checks(report_path, model):
    """逐行核对报表「附注」列。返回 NoteNoCheck 列表。"""
    wb = openpyxl.load_workbook(report_path)
    out = []

    # 每章一份 {归一化科目名: 编号}，以及章号
    chapter_info = {}
    for chapter in model.chapters:
        mapping = {}
        for sub in chapter.sub_items:
            if sub.number is None:
                continue
            mapping[rr.normalize_name(sub.name)] = sub.number
        chapter_info[chapter.scope] = (chapter, mapping)

    for sname in wb.sheetnames:
        if rr._kind_of(sname) is None:
            continue
        scope = rr._scope_of(sname)
        picked = chapter_info.get(scope)
        if picked is None and scope == '母公司':
            picked = chapter_info.get('单体')
        if picked is None:
            picked = next(iter(chapter_info.values()), None)
        if picked is None:
            continue
        chapter, mapping = picked
        if not mapping:
            continue

        ws = wb[sname]
        if not rn._sheet_has_data(ws):
            continue
        note_cols = rn._find_note_columns(ws)
        if not note_cols:
            continue

        for r in range(1, ws.max_row + 1):
            for nc in note_cols:
                if nc - 1 < 1:
                    continue
                cell_item = ws.cell(r, nc - 1)
                if isinstance(cell_item, MergedCell):
                    continue
                item_name = cell_item.value
                if not isinstance(item_name, str) or not item_name.strip():
                    continue
                number = rn._match_number(mapping, item_name)
                if number is None:
                    continue

                cell_note = ws.cell(r, nc)
                if isinstance(cell_note, MergedCell):
                    continue

                actual = str(cell_note.value or '').strip()
                expected = format_ref(chapter, number)
                chk = NoteNoCheck(
                    sheet=sname, row=r, item_name=item_name.strip(),
                    actual=actual, expected=expected, item_num=number,
                    chapter_num=chapter.number,
                )

                if not actual:
                    chk.status = 'fail'
                    chk.note = '附注号为空'
                    out.append(chk)
                    continue

                ach, aitem = parse_note_ref(actual)
                if ach is None or aitem is None:
                    chk.status = 'unmatched'
                    chk.note = '附注号写法无法识别'
                    out.append(chk)
                    continue

                if ach == chapter.number and aitem == number:
                    chk.status = 'pass'
                else:
                    chk.status = 'fail'
                    chk.note = f'章号或项目号不符（报表 {ach}-{aitem}，附注 {chapter.number}-{number}）'
                out.append(chk)
    return out


def summarize(checks):
    from collections import Counter
    return dict(Counter(c.status for c in checks))
