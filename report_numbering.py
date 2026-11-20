# -*- coding: utf-8 -*-
"""
report_numbering.py — 报表主表「附注」列的一键修复与补全。

做什么
------
报表（资产负债表 / 利润表 / 现金流量表）里通常有一列「附注」，写着该科目的
附注编号（如 `六、3`）。常见问题：
  · 空白没填
  · 编号与附注实际编号对不上（附注重排过、或换了年份模板）

本模块从 Word 附注解析出的子项目顺序/编号出发，把报表这列**补齐或改正**。

范围
----
合并 / 母公司分别处理：报表里叫「合并资产负债表」的用合并章的编号，
叫「母公司资产负债表」（或「资产负债表」）的用母公司章（没有则用单体章）的编号。

章节号不写死：用附注里识别出的真实章号（可能是「六、」，也可能是「七、」「十八、」）。
"""

import os
import re

import openpyxl
from openpyxl.cell.cell import MergedCell

import report_reader as rr

NOTE_HEADER_WORDS = ('附注', '附注号', '附注编号', '注释')

# 报表项目 → 附注子项目 的别名。
# 现金流量表里**只有「经营活动产生的现金流量净额」这一项**在附注里有对应小节
# （六、22 现金流量表补充资料）；其余如现金及现金等价物净增加额、期末/期初余额
# 都是内部小项或小计，不挂附注号。
REPORT_ITEM_ALIAS = {
    '经营活动产生的现金流量净额': '现金流量表补充资料',
    # 报表两行、附注一节
    '实收资本(或股本)净额': '实收资本(或股本)',
    '递延所得税资产': '递延所得税资产和递延所得税负债',
    '递延所得税负债': '递延所得税资产和递延所得税负债',
    # 利润表「其他综合收益的税后净额」及其子项，与附注 47 节同一编号
    '其他综合收益的税后净额': '归属于母公司所有者的其他综合收益',
    '权益法下不能转损益的其他综合收益': '归属于母公司所有者的其他综合收益',
    '其他权益工具投资公允价值变动': '归属于母公司所有者的其他综合收益',
}

# 这些报表项目**不需要**挂附注号：
#   · 营业总收入/营业总成本：本质是合计行，附注不单独披露
#   · 不能重分类进损益的其他综合收益：是「其他综合收益」的组成项
#   · 国家资本：实收资本的子项，在实收资本一节里披露
# 不需要编附注号的项目 = 合计项 + 母项的子项
# 统一由 report_item_terms 这张权威表判断，不再在这里零散枚举，
# 避免把「营业总成本」（合计项）和「营业成本」（真报表项目）混为一谈。
import report_item_terms as _terms

NO_NOTE_ITEMS = tuple(_terms.TOTALS) + tuple(_terms.SUB_ITEMS)


def is_no_note_item(name):
    """合计项、以及母项的子项，都不单独编附注号。"""
    if _terms.is_total(name):
        return True
    if _terms.parent_of(name) is not None:
        return True
    key = rr.normalize_name(name)
    return any(key == rr.normalize_name(w) for w in NO_NOTE_ITEMS)


def _chapter_prefix(chapter):
    """章号前缀，如「六、」。"""
    if chapter.number_text:
        text = chapter.number_text.strip()
        if text:
            return text
    if chapter.number is not None:
        return f'{chapter.number}、'
    return ''


def _pick_chapter(model, scope):
    """按范围挑章：合并→合并章；母公司→母公司章（没有就退回单体）。"""
    exact = [c for c in model.chapters if c.scope == scope]
    if exact:
        return exact[0]
    if scope == '母公司':
        single = [c for c in model.chapters if c.scope == '单体']
        if single:
            return single[0]
    return model.chapters[0] if model.chapters else None


def _find_note_columns(ws, max_scan=12):
    cols = []
    for r in range(1, min(max_scan, ws.max_row) + 1):
        for c in range(1, min(ws.max_column, 30) + 1):
            text = str(ws.cell(r, c).value or '').strip()
            if text in NOTE_HEADER_WORDS:
                cols.append(c)
    return sorted(set(cols))


def _strip_brackets(text):
    return re.sub(r'[\(（][^)）]*[)）]', '', text or '').strip()


def _match_number(items, item_name):
    """把报表科目名对到附注子项目，返回编号（int）或 None。

    只做**全等**匹配（允许去掉括号补充说明，如「实收资本（或：股本）」）。
    不做包含匹配——否则「长期应付职工薪酬」会被错配成「应付职工薪酬」的附注号，
    报错一个附注号比留空危害大得多。
    """
    key = rr.normalize_name(item_name)
    if not key:
        return None
    if key in items:
        return items[key]
    stripped = rr.normalize_name(_strip_brackets(item_name))
    if stripped and stripped in items:
        return items[stripped]
    # 别名：现金流量表的合计数挂到「现金流量表补充资料」那一节
    for alias_key, target in REPORT_ITEM_ALIAS.items():
        if key == alias_key or (stripped and stripped == alias_key):
            target_key = rr.normalize_name(target)
            if target_key in items:
                return items[target_key]
    return None


def _sheet_has_data(ws, max_scan=80):
    """表里有没有真实金额。空白模板表不需要填附注号。"""
    for r in range(1, min(max_scan, ws.max_row) + 1):
        for c in range(1, min(ws.max_column, 30) + 1):
            v = ws.cell(r, c).value
            if isinstance(v, (int, float)) and not isinstance(v, bool) and abs(v) > 0:
                return True
    return False


def fix_report_numbers(report_path, model, out_path=None, log=print):
    """按附注编号补全/修正报表的「附注」列，输出到新文件。

    返回 (输出路径, 变更说明列表)。
    """
    if not out_path:
        base, ext = os.path.splitext(report_path)
        out_path = f'{base}_附注编号已修复{ext}'
    if os.path.abspath(report_path) == os.path.abspath(out_path):
        raise ValueError('输出文件不能与原件相同，避免覆盖原始报表')

    wb = openpyxl.load_workbook(report_path)
    changes = []

    # 先按范围把每章的 {科目名: 编号} 建好
    items_by_scope = {}
    prefix_by_scope = {}
    for chapter in model.chapters:
        mapping = {}
        for sub in chapter.sub_items:
            if sub.number is None:
                continue
            mapping[rr.normalize_name(sub.name)] = sub.number
        items_by_scope[chapter.scope] = mapping
        prefix_by_scope[chapter.scope] = _chapter_prefix(chapter)

    for sname in wb.sheetnames:
        kind = rr._kind_of(sname)
        if kind is None:
            continue
        scope = rr._scope_of(sname)
        chapter = _pick_chapter(model, scope)
        if chapter is None:
            continue
        mapping = items_by_scope.get(scope) or items_by_scope.get(chapter.scope) or {}
        prefix = prefix_by_scope.get(scope) or prefix_by_scope.get(chapter.scope, '')
        if not mapping:
            continue

        ws = wb[sname]
        if not _sheet_has_data(ws):
            log(f'[编号] 跳过 {sname}：没有任何金额，是空白模板')
            continue
        note_cols = _find_note_columns(ws)
        if not note_cols:
            continue

        for r in range(1, ws.max_row + 1):
            for nc in note_cols:
                if nc - 1 < 1:
                    continue
                item_name = ws.cell(r, nc - 1).value
                if not isinstance(item_name, str) or not item_name.strip():
                    continue
                number = _match_number(mapping, item_name)
                if number is None:
                    continue
                target = f'{prefix}{number}'
                cell = ws.cell(r, nc)
                if isinstance(cell, MergedCell):
                    continue          # 合并区里的非左上角格子只读，跳过
                current = str(cell.value or '').strip()
                if current == target:
                    continue
                cell.value = target
                if current:
                    changes.append(
                        f'【{sname}】{item_name.strip()}（第{r}行）：'
                        f'原「{current}」→「{target}」')
                else:
                    changes.append(
                        f'【{sname}】{item_name.strip()}（第{r}行）：空白 → 「{target}」')

    wb.save(out_path)
    log(f'[编号] 报表附注号处理完成，共 {len(changes)} 处，已保存：{out_path}')
    return out_path, changes
