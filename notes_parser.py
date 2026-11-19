# -*- coding: utf-8 -*-
"""
notes_parser.py — 把 Word 附注解析成「章 → 子项目 → 表」的结构模型。

设计原则（按需求确定）
----------------------
1. **不认死章节编号**：附注编号可能是「六、」「七、」「（六）」「6、」，一律解出数字，
   下游要用编号时用解出来的值，不写死。
2. **合并与母公司都要**：凡标题命中「主要项目附注」类的章都收，并按标题判断范围
   （合并 / 母公司 / 单体）。母公司章名不统一（「母公司财务报表主要项目附注」
   或「公司财务报表项目附注」），所以按关键词判断而不是全等。
3. **模板只作参考**：本模块完全不读模板，纯靠文档自身结构识别；模板样板库是
   后续的可选加速器，不参与这里。
4. **识别不了要留痕**：所有"拿不准"的地方记进 doc.warnings，供成果表标黄提示。

对外接口
--------
    parse_notes(doc_path, work_dir=None) -> NotesDoc
    NotesDoc.chapters / .sub_items / .warnings / .summary()
"""

import os
import re
from dataclasses import dataclass, field
from typing import List, Optional

import docx_ir

# ---------------------------------------------------------------- 编号解析

CN_DIGITS = {'零': 0, '〇': 0, '一': 1, '二': 2, '两': 2, '三': 3, '四': 4,
             '五': 5, '六': 6, '七': 7, '八': 8, '九': 9}

# 「六、」「（六）」「第六章」等
RE_CHAPTER_NUM = re.compile(
    r'^\s*(?:第\s*)?[（(]?\s*([一二三四五六七八九十百零〇两]+|\d+)\s*[)）]?\s*'
    r'(?:章|、|．|\.)\s*'
)


def cn_to_int(text):
    """中文数字 → 整数。支持 六 / 十二 / 二十三 / 一百。失败返回 None。"""
    if not text:
        return None
    s = text.strip()
    if s.isdigit():
        return int(s)
    if '百' in s:
        head, _, tail = s.partition('百')
        h = CN_DIGITS.get(head, 1) if head else 1
        rest = cn_to_int(tail) if tail else 0
        return h * 100 + (rest or 0)
    if '十' in s:
        head, _, tail = s.partition('十')
        h = CN_DIGITS.get(head, 1) if head else 1
        t = CN_DIGITS.get(tail, 0) if tail else 0
        return h * 10 + t
    # 纯个位串，如 一二三（少见，逐位累加）
    total = 0
    for ch in s:
        if ch not in CN_DIGITS:
            return None
        total = total * 10 + CN_DIGITS[ch]
    return total or None


def parse_heading_number(text):
    """从标题里解出编号，返回 (number, 去掉编号后的标题)。"""
    m = RE_CHAPTER_NUM.match(text or '')
    if not m:
        return None, (text or '').strip()
    num = cn_to_int(m.group(1))
    rest = text[m.end():].strip()
    return num, rest


# ---------------------------------------------------------------- 数据模型

@dataclass
class TableRef:
    """一张附注表在文档流里的位置。"""
    idx: int                      # 表格序号（文档内 0 基）
    rows: int
    cols: int
    grid: list                    # 网格，每格 {'t': 文本, 'm': 合并标记}
    stream_pos: int = -1          # 在 stream 里的下标

    @property
    def nonempty_cells(self):
        return sum(1 for row in self.grid for c in row if c['m'] == ' ' and c['t'])


@dataclass
class SubItem:
    """一个附注子项目，例如「3、预付款项」。"""
    seq: int                      # 章内出现顺序（1 基，最可靠）
    name: str                     # 去掉编号后的科目名
    raw_title: str                # 原始标题文字
    number: Optional[int]         # 编号：优先取正文手打的，其次由 Word 自动编号还原
    heading_stream_pos: int
    number_text: Optional[str] = None    # 编号的显示形式，如 "3、" / "六、"
    tables: List[TableRef] = field(default_factory=list)
    paragraphs: List[str] = field(default_factory=list)   # 表间的小标题/说明文字

    @property
    def has_number(self):
        return self.number is not None


@dataclass
class Chapter:
    """一个「主要项目附注」章。"""
    name: str                     # 原始标题
    number: Optional[int]         # 章编号（六 → 6），手打优先，其次自动编号还原
    scope: str                    # 合并 / 母公司 / 单体
    outline: int
    start: int
    end: int
    number_text: Optional[str] = None
    sub_items: List[SubItem] = field(default_factory=list)

    @property
    def table_count(self):
        return sum(1 for s in self.sub_items for _ in s.tables)


@dataclass
class NotesDoc:
    path: str
    docx_path: str
    stream: list
    chapters: List[Chapter]
    warnings: List[str] = field(default_factory=list)

    @property
    def sub_items(self):
        return [s for c in self.chapters for s in c.sub_items]

    def summary(self):
        lines = [f'源文件: {os.path.basename(self.path)}',
                 f'段落 {sum(1 for x in self.stream if x["kind"] == "p")} 个 / '
                 f'表格 {sum(1 for x in self.stream if x["kind"] == "table")} 张',
                 f'识别到主要项目附注章 {len(self.chapters)} 个']
        for ch in self.chapters:
            if ch.number_text:
                num = f'编号「{ch.number_text}」'
            elif ch.number is not None:
                num = f'编号 {ch.number}'
            else:
                num = '无编号'
            lines.append(f'  章<{ch.name}>  {num}  范围={ch.scope}  '
                         f'子项目 {len(ch.sub_items)} 个  表 {ch.table_count} 张')
            for s in ch.sub_items:
                if s.number_text:
                    flag = s.number_text.strip()
                elif s.has_number:
                    flag = f'#{s.number}'
                else:
                    flag = f'seq{s.seq}'
                titles = '; '.join(s.paragraphs[:2])
                if len(s.paragraphs) > 2:
                    titles += ' …'
                lines.append(f'    {flag:<8} {s.name:<24} 表{len(s.tables)} '
                             f'{"| " + titles[:40] if titles else ""}')
        if self.warnings:
            lines.append('需要注意：')
            lines.extend('  - ' + w for w in self.warnings)
        return '\n'.join(lines)


# ---------------------------------------------------------------- 主流程

def parse_notes(doc_path, work_dir=None, stream=None):
    """解析附注文档，返回 NotesDoc。"""
    docx_path = docx_ir.ensure_docx(doc_path, work_dir=work_dir)
    if stream is None:
        stream = docx_ir.load_stream(docx_path)

    warnings = []
    raw_chapters = docx_ir.find_note_chapters(stream)

    if not raw_chapters:
        raise ValueError(
            '未在文档中找到「主要项目附注」章节。\n'
            '识别关键词：' + '、'.join(docx_ir.NOTE_CHAPTER_KEYS) + '\n'
            '请确认该章节标题，或确认文档标题用了大纲级别（Word 的「标题 1/2」样式）。'
        )

    chapters = []
    for rc in raw_chapters:
        number, rest = parse_heading_number(rc['name'])
        head = stream[rc['start']]
        # 正文没手打编号，就用 Word 自动编号还原出来的值
        if number is None:
            number = head.get('auto_number_value')
            number_text = head.get('auto_number')
        else:
            # 手打编号：把「六、」这一段原文留下来，用于生成报表附注号
            m = RE_CHAPTER_NUM.match(rc['name'])
            number_text = m.group(0).strip() if m else head.get('auto_number')
        ch = Chapter(
            name=rc['name'],
            number=number,
            number_text=number_text,
            scope=rc['scope'],
            outline=rc['outline'],
            start=rc['start'],
            end=rc['end'],
        )

        raw_items = docx_ir.find_sub_items(stream, rc)
        if not raw_items:
            warnings.append(f'章<{rc["name"]}> 下没识别出子项目，可能需要检查标题级别')

        for ri in raw_items:
            item_head = stream[ri['idx']]
            num = ri['number']
            if num is None:
                num = item_head.get('auto_number_value')
            sub = SubItem(
                seq=ri['seq'],
                name=ri['name'],
                raw_title=ri['raw_title'],
                number=num,
                number_text=item_head.get('auto_number'),
                heading_stream_pos=ri['idx'],
            )
            # 收集该子项目区间内的表与说明段落
            for pos in range(ri['start'], ri['end']):
                it = stream[pos]
                if it['kind'] == 'table':
                    sub.tables.append(TableRef(
                        idx=it['idx'], rows=it['rows'], cols=it['cols'],
                        grid=it['grid'], stream_pos=pos,
                    ))
                elif it['kind'] == 'p' and it['text']:
                    if it.get('outline') is None:
                        sub.paragraphs.append(it['text'])
            if not sub.tables:
                warnings.append(
                    f'章<{rc["name"]}> 子项目「{sub.name}」下没有表格'
                    f'（可能只写了文字说明，或缺表）'
                )
            ch.sub_items.append(sub)

        # 编号连续性体检（只在有编号时才有意义）
        numbers = [s.number for s in ch.sub_items if s.has_number]
        if len(numbers) >= 2:
            expected = list(range(numbers[0], numbers[0] + len(numbers)))
            if numbers != expected:
                warnings.append(
                    f'章<{rc["name"]}> 子项目编号不连续：'
                    f'实际 {numbers[:8]}{"…" if len(numbers) > 8 else ""}。'
                    f'成果表按出现顺序统一编号。'
                )
        elif numbers and len(numbers) != len(ch.sub_items):
            warnings.append(
                f'章<{rc["name"]}> 部分子项目没有编号（{len(ch.sub_items) - len(numbers)} 个），'
                f'成果表按出现顺序统一编号。'
            )

        chapters.append(ch)

    return NotesDoc(path=os.path.abspath(doc_path), docx_path=docx_path,
                    stream=stream, chapters=chapters, warnings=warnings)


# ---------------------------------------------------------------- CLI

def _main():
    import sys
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    target = sys.argv[1]
    if os.path.isdir(target):
        files = [os.path.join(target, f) for f in sorted(os.listdir(target))
                 if f.lower().endswith(('.doc', '.docx')) and not f.startswith('~$')]
    else:
        files = [target]

    for f in files:
        print('=' * 90)
        try:
            doc = parse_notes(f, work_dir='_work/parse')
            print(doc.summary())
        except Exception as exc:
            print(f'解析失败 {f}: {exc}')
    return 0


if __name__ == '__main__':
    raise SystemExit(_main())
