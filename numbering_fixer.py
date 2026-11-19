# -*- coding: utf-8 -*-
"""
numbering_fixer.py — 附注子项目编号的体检与重建。

为什么需要
----------
Word 里的编号有三种情况，还可能叠加：
    ① 只有人工编号：正文里就写着「1、货币资金」
    ② 只有自动编号：段落挂 numId，正文里没有数字
    ③ 两者都有：Word 显示成「1、1、货币资金」
再加上自动编号会重启（1–8 之后又回到 1），单靠解析只能"看出问题"，
不能根治。

本模块做两件事
--------------
1. **体检** diagnose()：逐个子项目给出
      当前自动编号 / 当前手工编号 / 实际显示 / 问题类型 / 建议编号
   只读，不碰文件。

2. **重建** renumber_chapter()：按需求把该章子项目统一改成**自动编号**
       - 为该章新建一个专用编号列表（从 1 开始、十进制、格式「%1、」）
       - 整章子项目共用同一个列表 —— 共用才能连续，不会被重启打断
       - 删掉段落里手打的编号文字
       - 给标题段挂上新列表
       - 章标题、正文、表格一律不动
       - 缩进/悬挂设为 0、suff=nothing，编号紧贴文字，不改排版
   永远写到新文件，不覆盖原件；写完立刻重新解析复核 1..N 是否连续。

注意：Word 的自动编号值**无法按段落单独指定**（只能靠列表重启控制）。
所以"1–8 之后又重启成 1"这种错乱，唯一干净的解是把该章所有子项目
收编到同一个新列表里，让 Word 重新从 1 连续算。
"""

import os
import re
import shutil
from dataclasses import dataclass
from typing import List, Optional

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

# 手工编号的几种常见写法
RE_MANUAL_PREFIX = re.compile(
    r'^\s*(?:'
    r'\d+\s*[、．.](?!\d)'         # 1、 1. 1．
    r'|（\s*\d+\s*）'              # （1）
    r'|\(\s*\d+\s*\)'              # (1)
    r'|[一二三四五六七八九十百]+\s*[、．.]'   # 一、
    r'|[①-⑳]'                     # ①
    r')\s*'
)


# ---------------------------------------------------------------- 数据模型

@dataclass
class NumIssue:
    chapter: str
    seq: int                 # 章内出现顺序（1 基）
    name: str
    auto_label: Optional[str]        # Word 自动编号
    manual_prefix: Optional[str]     # 正文里手打的编号
    display: str                     # 实际显示（自动+手工拼起来）
    problem: str                     # 正常 / 双编号 / 重启 / 重号 / 跳号 / 无编号
    target_text: str                 # 建议改成

    def line(self):
        return (f'{self.seq:>3}  {self.name[:22]:<24} '
                f'{(self.auto_label or "—"):<6} {(self.manual_prefix or "—"):<6} '
                f'{self.display or "（空）":<8} {self.problem:<6} {self.target_text}')


# ---------------------------------------------------------------- 体检

def _manual_prefix_of(paragraph):
    """返回 (前缀文字, 前缀长度) 或 (None, 0)。"""
    return _manual_prefix_of_text(paragraph.text)


def _manual_prefix_of_text(raw):
    raw = raw or ''
    m = RE_MANUAL_PREFIX.match(raw)
    if not m:
        return None, 0
    return m.group(0), len(m.group(0))


def diagnose(doc_model):
    """对 NotesDoc 的所有章做编号体检，返回 List[NumIssue]。"""
    issues = []
    for ch in doc_model.chapters:
        prev_values = []
        seen = {}
        for sub in ch.sub_items:
            head = doc_model.stream[sub.heading_stream_pos]
            auto = head.get('auto_number')
            manual, _ = _manual_prefix_of_text(head.get('text'))

            display = ''.join(x for x in (auto, manual) if x)
            v = sub.number if sub.number is not None else sub.seq

            if auto and manual:
                problem = '双编号'
            elif not auto and not manual:
                problem = '无编号'
            elif v in seen:
                problem = '重号'
            elif prev_values and v < prev_values[-1]:
                problem = '重启' if v == 1 else '倒序'
            elif v != sub.seq:
                problem = '跳号'
            else:
                problem = '正常'

            issues.append(NumIssue(
                chapter=ch.name,
                seq=sub.seq,
                name=sub.name,
                auto_label=auto,
                manual_prefix=manual,
                display=display,
                problem=problem,
                target_text=f'{sub.seq}、',
            ))
            prev_values.append(v)
            seen[v] = True
    return issues


# ---------------------------------------------------------------- 重建

def _next_ids(numbering_root):
    """取下一个可用的 abstractNumId / numId。"""
    max_a = -1
    for an in numbering_root.findall(qn('w:abstractNum')):
        try:
            max_a = max(max_a, int(an.get(qn('w:abstractNumId'))))
        except (TypeError, ValueError):
            pass
    max_n = 0
    for num in numbering_root.findall(qn('w:num')):
        try:
            max_n = max(max_n, int(num.get(qn('w:numId'))))
        except (TypeError, ValueError):
            pass
    return max_a + 1, max_n + 1


def _build_single_level_list(numbering_root, lvl_text='%1、'):
    """新建一个单级、从 1 开始、编号紧贴文字的列表定义。返回 numId。"""
    aid, nid = _next_ids(numbering_root)

    abstract = OxmlElement('w:abstractNum')
    abstract.set(qn('w:abstractNumId'), str(aid))

    multi = OxmlElement('w:multiLevelType')
    multi.set(qn('w:val'), 'singleLevel')
    abstract.append(multi)

    lvl = OxmlElement('w:lvl')
    lvl.set(qn('w:ilvl'), '0')

    for tag, val in (('w:start', '1'), ('w:numFmt', 'decimal'),
                     ('w:lvlText', lvl_text), ('w:lvlJc', 'left')):
        el = OxmlElement(tag)
        el.set(qn('w:val'), val)
        lvl.append(el)

    pPr = OxmlElement('w:pPr')
    ind = OxmlElement('w:ind')
    # 缩进与悬挂都归零：编号紧贴左边，不改变原有排版位置
    ind.set(qn('w:left'), '0')
    ind.set(qn('w:hanging'), '0')
    pPr.append(ind)
    # 编号与正文之间不插入制表符/空格（lvlText 自带「、」）
    suff = OxmlElement('w:suff')
    suff.set(qn('w:val'), 'nothing')
    pPr.append(suff)
    lvl.append(pPr)

    abstract.append(lvl)

    # 架构要求：所有 abstractNum 必须排在所有 num 之前
    first_num = numbering_root.find(qn('w:num'))
    if first_num is not None:
        first_num.addprevious(abstract)
    else:
        numbering_root.append(abstract)

    num = OxmlElement('w:num')
    num.set(qn('w:numId'), str(nid))
    ref = OxmlElement('w:abstractNumId')
    ref.set(qn('w:val'), str(aid))
    num.append(ref)
    numbering_root.append(num)

    return nid


def _set_numpr(paragraph, num_id, ilvl=0):
    """给段落设置 numPr（覆盖原有自动编号）。"""
    pPr = paragraph._p.get_or_add_pPr()
    numPr = pPr.find(qn('w:numPr'))
    if numPr is None:
        numPr = OxmlElement('w:numPr')
        # pPr 的子元素有固定顺序，numPr 紧跟在 pStyle/keepNext 等之后；
        # 这里用 python-docx 的插入器保证位置合法。
        try:
            pPr.get_or_add_numPr()
            numPr = pPr.find(qn('w:numPr'))
        except AttributeError:
            pPr.append(numPr)

    ilvl_el = numPr.find(qn('w:ilvl'))
    if ilvl_el is None:
        ilvl_el = OxmlElement('w:ilvl')
        numPr.append(ilvl_el)
    ilvl_el.set(qn('w:val'), str(ilvl))

    numid_el = numPr.find(qn('w:numId'))
    if numid_el is None:
        numid_el = OxmlElement('w:numId')
        numPr.append(numid_el)
    numid_el.set(qn('w:val'), str(num_id))


def _strip_prefix_from_runs(paragraph, prefix_len):
    """把段落开头 prefix_len 个字符（手工编号）从各个 run 里删掉，保留其余格式。"""
    remaining = prefix_len
    for run in paragraph.runs:
        if remaining <= 0:
            break
        text = run.text
        if not text:
            continue
        if len(text) <= remaining:
            remaining -= len(text)
            run.text = ''
        else:
            run.text = text[remaining:]
            remaining = 0


def _body_paragraphs(doc):
    """按文档顺序返回正文顶层段落对象列表（与 docx_ir.load_stream 的 p_idx 对齐）。"""
    out = []
    for child in doc.element.body.iterchildren():
        if child.tag == qn('w:p'):
            out.append(Paragraph(child, doc))
    return out


def _apply_chapter(doc, numbering_root, doc_model, chapter, log):
    """在一个已打开的 doc 上，把该章子项目收编到新建的专用列表。返回处理段数。"""
    num_id = _build_single_level_list(numbering_root)
    log(f'[编号] 章<{chapter.name}> 新建列表 numId={num_id}，从 1 开始')

    paras = _body_paragraphs(doc)
    changed = 0
    for sub in chapter.sub_items:
        p_index = doc_model.stream[sub.heading_stream_pos].get('i')
        if p_index is None or p_index >= len(paras):
            log(f'[编号] 跳过（找不到段落）：{sub.name}')
            continue
        para = paras[p_index]
        _, plen = _manual_prefix_of(para)
        if plen:
            _strip_prefix_from_runs(para, plen)
        _set_numpr(para, num_id, ilvl=0)
        changed += 1
    return changed


def renumber_document(doc_model, out_path, chapters=None, log=print):
    """把所有（或指定）主要项目附注章的子项目改成连续自动编号，输出一个文件。

    每章各自新建一个专用列表，所以每章都从 1 重新开始连续。
    章标题、正文、表格一律不动。

    返回 (out_path, issues, verify_detail)，verify_detail 为 {章名: (是否连续, 实际值)}
    """
    src = doc_model.path
    if os.path.abspath(src) == os.path.abspath(out_path):
        raise ValueError('输出文件不能与原件相同，避免覆盖原始附注')
    if not out_path.lower().endswith('.docx'):
        raise ValueError('重建编号请输出 .docx（.doc 需先经 Word 转换）')
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or '.', exist_ok=True)

    # .doc 无法用 python-docx 直接打开：先归一成 docx 再复制
    import docx_ir
    src_docx = docx_ir.ensure_docx(src, work_dir=os.path.join('_work', 'parse'))
    shutil.copy2(src_docx, out_path)

    targets = chapters if chapters is not None else doc_model.chapters
    wanted = {ch.name for ch in targets}
    issues = [i for i in diagnose(doc_model) if i.chapter in wanted]

    doc = Document(out_path)
    try:
        numbering_root = doc.part.numbering_part.element
    except (KeyError, NotImplementedError, AttributeError):
        raise RuntimeError('该文档没有编号定义(numbering.xml)，无法重建自动编号')

    total = 0
    for ch in targets:
        total += _apply_chapter(doc, numbering_root, doc_model, ch, log)

    doc.save(out_path)
    log(f'[编号] 共处理 {total} 个子项目，保存至 {out_path}')

    # 立刻复核：重新解析，确认每章 1..N 连续
    import notes_parser
    check = notes_parser.parse_notes(out_path, work_dir=os.path.dirname(os.path.abspath(out_path)))
    verify_detail = {}
    for ch in targets:
        found = next((c for c in check.chapters if c.name == ch.name), None)
        if found is None and len(check.chapters) == 1:
            found = check.chapters[0]
        if found is None:
            verify_detail[ch.name] = (False, [])
            continue
        values = [s.number for s in found.sub_items]
        ok = values == list(range(1, len(values) + 1))
        verify_detail[ch.name] = (ok, values)
        shown = values[:12]
        log(f'[编号] 复核<{ch.name}>：{shown}{"…" if len(values) > 12 else ""} '
            f'{"连续 ✓" if ok else "不连续 ✗ 期望 " + str(list(range(1, len(values) + 1))[:12])}')
    return out_path, issues, verify_detail


def renumber_chapter(doc_model, chapter, out_path, log=print):
    """只处理单章（内部用）。日常请用 renumber_document。"""
    return renumber_document(doc_model, out_path, chapters=[chapter], log=log)


# ---------------------------------------------------------------- CLI

def _main():
    import sys
    import notes_parser
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    path = sys.argv[1]
    apply = '--apply' in sys.argv

    model = notes_parser.parse_notes(path, work_dir='_work/parse')
    issues = diagnose(model)
    print(f"{'序号':>3}  {'子项目':<24} {'自动':<6} {'手工':<6} {'现在显示':<8} {'问题':<6} 改成")
    print('-' * 96)
    for it in issues:
        print(it.line())

    if not apply:
        print('\n（只体检，未修改。加 --apply 才会重建编号）')
        return 0

    base = os.path.splitext(path)[0]
    out = f'{base}_编号已规范.docx'
    _, _, detail = renumber_document(model, out)
    all_ok = all(ok for ok, _ in detail.values())
    print('\n复核结果：' + ('全部连续 ✓' if all_ok else '存在不连续 ✗，请检查输出文件'))
    return 0 if all_ok else 1


if __name__ == '__main__':
    raise SystemExit(_main())
