# -*- coding: utf-8 -*-
"""
probe_numbering.py — 探查 Word 自动编号（numbering.xml）能不能还原出章节号。

Word 的自动编号不写在正文文本里，而是：
    段落 <w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr>
    numbering.xml: <w:num w:numId="3"><w:abstractNumId w:val="7"/></w:num>
                   <w:abstractNum w:abstractNumId="7">
                     <w:lvl w:ilvl="0"><w:start w:val="1"/>
                       <w:numFmt w:val="chineseCounting"/> / upperRoman / decimal ...
                       <w:lvlText w:val="%1、"/></w:lvl>
真实显示的编号 = 该 (numId, ilvl) 组合下，本文档中此前出现过几次 + start 值。

用法：
    python tools/probe_numbering.py "_work/docx/一般企业模板6.4财务报表附注.docx"
"""

import sys

from docx import Document
from docx.oxml.ns import qn


def q(el, tag):
    return el.find(qn(tag)) if el is not None else None


def load_numbering(doc):
    """numId -> abstractNumId -> {ilvl: {start, numFmt, lvlText}}"""
    try:
        part = doc.part.numbering_part
    except (KeyError, NotImplementedError, AttributeError):
        return {}, {}
    root = part.element
    abstract = {}
    for an in root.findall(qn('w:abstractNum')):
        aid = an.get(qn('w:abstractNumId'))
        lvls = {}
        for lvl in an.findall(qn('w:lvl')):
            il = int(lvl.get(qn('w:ilvl')))
            st = q(lvl, 'w:start')
            fmt = q(lvl, 'w:numFmt')
            txt = q(lvl, 'w:lvlText')
            lvls[il] = {
                'start': int(st.get(qn('w:val'))) if st is not None else 1,
                'numFmt': fmt.get(qn('w:val')) if fmt is not None else 'decimal',
                'lvlText': txt.get(qn('w:val')) if txt is not None else '',
            }
        abstract[aid] = lvls

    nums = {}
    for num in root.findall(qn('w:num')):
        nid = num.get(qn('w:numId'))
        ref = q(num, 'w:abstractNumId')
        if ref is not None:
            nums[nid] = ref.get(qn('w:val'))
    return abstract, nums


def para_numpr(p):
    pPr = q(p._p, 'w:pPr')
    if pPr is None:
        return None
    numPr = q(pPr, 'w:numPr')
    if numPr is None:
        return None
    ilvl = q(numPr, 'w:ilvl')
    numId = q(numPr, 'w:numId')
    return {
        'ilvl': int(ilvl.get(qn('w:val'))) if ilvl is not None else 0,
        'numId': numId.get(qn('w:val')) if numId is not None else None,
    }


CN = ['零', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十']


def cn_num(n):
    if n <= 10:
        return CN[n]
    if n < 20:
        return '十' + (CN[n - 10] if n > 10 else '')
    if n < 100:
        t, d = divmod(n, 10)
        return CN[t] + '十' + (CN[d] if d else '')
    return str(n)


def fmt_number(value, numfmt):
    if numfmt in ('chineseCounting', 'chineseCountingThousand', 'taiwaneseCounting'):
        return cn_num(value)
    if numfmt == 'upperRoman':
        vals = [(1000, 'M'), (900, 'CM'), (500, 'D'), (400, 'CD'), (100, 'C'),
                (90, 'XC'), (50, 'L'), (40, 'XL'), (10, 'X'), (9, 'IX'),
                (5, 'V'), (4, 'IV'), (1, 'I')]
        out = ''
        for v, s in vals:
            while value >= v:
                out += s
                value -= v
        return out
    if numfmt == 'lowerLetter':
        return chr(ord('a') + value - 1)
    if numfmt == 'upperLetter':
        return chr(ord('A') + value - 1)
    return str(value)


def main():
    path = sys.argv[1]
    doc = Document(path)
    abstract, nums = load_numbering(doc)
    print(f'numbering: abstractNum={len(abstract)}  num={len(nums)}')

    counters = {}      # (numId, ilvl) -> 已出现次数
    found = 0
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        np = para_numpr(p)
        if not np or not np['numId']:
            continue
        key = (np['numId'], np['ilvl'])
        counters[key] = counters.get(key, 0) + 1
        aid = nums.get(np['numId'])
        lvl = abstract.get(aid, {}).get(np['ilvl'])
        if not lvl:
            continue
        value = lvl['start'] + counters[key] - 1
        shown = fmt_number(value, lvl['numFmt'])
        label = lvl['lvlText'].replace(f'%{np["ilvl"] + 1}', shown)
        for i in range(1, 10):
            label = label.replace(f'%{i}', '')
        if found < 60:
            print(f'  numId={np["numId"]:<4} ilvl={np["ilvl"]} '
                  f'fmt={lvl["numFmt"]:<22} 显示=<{label}>  {text[:44]}')
        found += 1
    print(f'带自动编号的段落共 {found} 个')


if __name__ == '__main__':
    main()
