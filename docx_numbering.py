# -*- coding: utf-8 -*-
"""
docx_numbering.py — 还原 Word 自动编号，算出段落真实显示的编号文字。

背景
----
附注的章节号通常是 Word 自动编号，**正文文本里根本没有「六、」这两个字**。
编号信息分散在三处：

    <w:p><w:pPr><w:numPr>
            <w:ilvl w:val="0"/>
            <w:numId w:val="1"/>
        </w:numPr></w:pPr></w:p>

    numbering.xml:
        <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
        <w:abstractNum w:abstractNumId="0">
          <w:lvl w:ilvl="0">
            <w:start w:val="1"/>
            <w:numFmt w:val="japaneseCounting"/>
            <w:lvlText w:val="%1、"/>
          </w:lvl>
        </w:abstractNum>

真实编号 = 该 (numId, ilvl) 在本段落之前的出现次数 + start。
要点：**上级推进时，下级要重置**（否则「编制基础」会被数成 5、而不是 1、）。
另外 <w:lvlRestart w:val="0"/> 表示该级不随上级重置，也需要照顾。

对外接口
--------
    resolver = NumberingResolver(doc)     # doc 是 python-docx 的 Document
    resolver.lookup(paragraph)            # -> {'label','value','ilvl'} 或 None
    walk_labels(doc, paragraphs)          # 按顺序一次性算完（推荐，性能好）
"""

import re

from docx.oxml.ns import qn

CN = '零一二三四五六七八九'


def _q(el, tag):
    return el.find(qn(tag)) if el is not None else None


# ---------------------------------------------------------------- 数字格式

def _cn(n):
    """整数 → 中文数字（1 → 一，12 → 十二，105 → 一百零五）。"""
    if n <= 0:
        return str(n)
    if n < 10:
        return CN[n]
    if n < 20:
        return '十' + (CN[n % 10] if n % 10 else '')
    if n < 100:
        t, d = divmod(n, 10)
        return CN[t] + '十' + (CN[d] if d else '')
    if n < 1000:
        h, r = divmod(n, 100)
        out = CN[h] + '百'
        if r:
            out += ('零' + CN[r] if r < 10 else _cn(r))
        return out
    return str(n)


CIRCLED = '①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳'
PAREN = '⑴⑵⑶⑷⑸⑹⑺⑻⑼⑽⑾⑿⒀⒁⒂⒃⒄⒅⒆⒇'


def _roman(n):
    if n <= 0:
        return str(n)
    vals = [(1000, 'M'), (900, 'CM'), (500, 'D'), (400, 'CD'), (100, 'C'),
            (90, 'XC'), (50, 'L'), (40, 'XL'), (10, 'X'), (9, 'IX'),
            (5, 'V'), (4, 'IV'), (1, 'I')]
    out = ''
    for v, s in vals:
        while n >= v:
            out += s
            n -= v
    return out


def format_number(n, fmt):
    """按 Word 的 numFmt 把整数转成显示文字。"""
    if not fmt:
        return str(n)
    if fmt in ('chineseCounting', 'chineseCountingThousand', 'japaneseCounting',
               'taiwaneseCounting', 'taiwaneseCountingThousand', 'ideographDigital'):
        return _cn(n)
    if fmt == 'upperRoman':
        return _roman(n)
    if fmt == 'lowerRoman':
        return _roman(n).lower()
    if fmt == 'decimalEnclosedCircle':
        return CIRCLED[n - 1] if 1 <= n <= len(CIRCLED) else str(n)
    if fmt == 'decimalEnclosedParen':
        return PAREN[n - 1] if 1 <= n <= len(PAREN) else str(n)
    if fmt == 'upperLetter':
        return chr(ord('A') + n - 1) if n >= 1 else str(n)
    if fmt == 'lowerLetter':
        return chr(ord('a') + n - 1) if n >= 1 else str(n)
    if fmt == 'none':
        return ''
    return str(n)


# ---------------------------------------------------------------- 定义表

def load_definitions(doc):
    """返回 (abstract, nums)：
    abstract: abstractNumId -> {ilvl: {start, numFmt, lvlText, restart}}
    nums:     numId -> abstractNumId
    """
    try:
        root = doc.part.numbering_part.element
    except (KeyError, NotImplementedError, AttributeError):
        return {}, {}

    abstract = {}
    for an in root.findall(qn('w:abstractNum')):
        aid = an.get(qn('w:abstractNumId'))
        lvls = {}
        for lvl in an.findall(qn('w:lvl')):
            try:
                il = int(lvl.get(qn('w:ilvl')))
            except (TypeError, ValueError):
                continue
            st, fmt, txt, rs = (_q(lvl, 'w:start'), _q(lvl, 'w:numFmt'),
                                _q(lvl, 'w:lvlText'), _q(lvl, 'w:lvlRestart'))
            lvls[il] = {
                'start': int(st.get(qn('w:val'))) if st is not None else 1,
                'numFmt': fmt.get(qn('w:val')) if fmt is not None else 'decimal',
                'lvlText': txt.get(qn('w:val')) if txt is not None else '',
                # restart=None 表示默认随上级重置；0 表示永不重置
                'restart': int(rs.get(qn('w:val'))) if rs is not None else None,
            }
        abstract[aid] = lvls

    nums = {}
    for num in root.findall(qn('w:num')):
        nid = num.get(qn('w:numId'))
        ref = _q(num, 'w:abstractNumId')
        if ref is not None:
            nums[nid] = ref.get(qn('w:val'))
    return abstract, nums


def get_numpr(paragraph):
    """取段落的 (numId, ilvl)。段落没挂就沿样式链往上找。"""
    def from_ppr(pPr):
        if pPr is None:
            return None
        numPr = _q(pPr, 'w:numPr')
        if numPr is None:
            return None
        ilvl = _q(numPr, 'w:ilvl')
        numId = _q(numPr, 'w:numId')
        return {
            'numId': numId.get(qn('w:val')) if numId is not None else None,
            'ilvl': int(ilvl.get(qn('w:val'))) if ilvl is not None else 0,
        }

    got = from_ppr(_q(paragraph._p, 'w:pPr'))
    if got and got.get('numId'):
        return got

    style = paragraph.style
    for _ in range(10):
        if style is None:
            break
        got = from_ppr(style.element.find(qn('w:pPr')))
        if got and got.get('numId'):
            return got
        style = style.base_style
    return None


# ---------------------------------------------------------------- 计数

class NumberingResolver:
    """按文档顺序喂段落，算出每段真实显示的编号。"""

    def __init__(self, doc):
        self.abstract, self.nums = load_definitions(doc)
        self.counters = {}          # numId -> {ilvl: 当前值}

    def _value(self, num_id, ilvl, defs):
        c = self.counters.setdefault(num_id, {})
        start = defs.get(ilvl, {}).get('start', 1)
        c.setdefault(ilvl, start - 1)
        c[ilvl] += 1

        # 上级推进时重置下级（restart == 0 的级别除外）
        for deeper in [k for k in list(c) if k > ilvl]:
            if defs.get(deeper, {}).get('restart') == 0:
                continue
            c[deeper] = defs.get(deeper, {}).get('start', 1) - 1
        return c[ilvl]

    def label(self, paragraph):
        """返回 {'label': '六、', 'value': 6, 'ilvl': 0} 或 None。"""
        np = get_numpr(paragraph)
        if not np or not np.get('numId'):
            return None
        num_id = np['numId']
        ilvl = np['ilvl']
        aid = self.nums.get(num_id)
        defs = self.abstract.get(aid, {})
        if not defs:
            return None

        value = self._value(num_id, ilvl, defs)
        lvl = defs.get(ilvl, {})
        text = lvl.get('lvlText') or ''
        c = self.counters.get(num_id, {})

        # lvlText 里可能引用多级：%1、%2 ...
        def repl(m):
            k = int(m.group(1))
            v = c.get(k - 1, defs.get(k - 1, {}).get('start', 1))
            return format_number(v, defs.get(k - 1, {}).get('numFmt'))

        label = re.sub(r'%(\d)', repl, text)
        return {'label': label, 'value': value, 'ilvl': ilvl,
                'numFmt': lvl.get('numFmt')}


def walk_labels(doc):
    """按文档顺序返回 {段落元素: 编号信息}，供 load_stream 使用。

    只对正文顶层段落编号；表格里的段落不参与（附注表格内一般不带章节编号）。
    """
    from docx.text.paragraph import Paragraph
    resolver = NumberingResolver(doc)
    out = {}
    for child in doc.element.body.iterchildren():
        if child.tag != qn('w:p'):
            continue
        p = Paragraph(child, doc)
        info = resolver.label(p)
        if info:
            out[child] = info
    return out
