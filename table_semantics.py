# -*- coding: utf-8 -*-
"""
table_semantics.py — 让程序"看懂"附注表的列名和行名。

解决的问题
----------
附注表的表头通常是 1~3 行、并且带合并单元格：

    第1行: 账龄 │ 期末余额（跨3列）            │ 期初余额（跨3列）
    第2行:      │ 账面余额 │ 坏账准备 │        │ 账面余额 │ 坏账准备
    第3行:      │ 金额 │ 比例 │ 坏账准备 │      │ 金额 │ 比例 │ 坏账准备

只看"第 2 列"没有任何意义；要把它还原成一条路径
「期末余额 → 账面余额 → 金额」才知道这列是什么，进而才能判断：
    · 账面价值 = 账面余额 − 坏账准备
    · 合计 = 上面若干明细行之和
    · 这一列对应报表的哪一期

合并单元格的处理
----------------
docx_ir 给出的网格里，被横向合并覆盖的格子标记为 'h'/'x'、纵向为 'v'。
本模块在拼表头时把 'h'/'x' 向左继承、'v' 向上继承，
所以"期末余额"能正确铺满它跨的 3 列。

期间的解析
----------
模板表头常写「期末余额 / 期初余额」而不是具体日期，真实附注多写具体日期。
两种都支持：
    1. 表头里直接写日期 → 直接用
    2. 表头写「期末/期初/本期/上期」→ 从本章正文里找对照句，例如
       「…"期初"指2025年1月1日，"期末"指2025年12月31日，"本期"指2025年度…」
       解析出来再用；找不到就留空，交给上层标"待确认"。

对外接口
--------
    ctx   = build_chapter_context(doc_model, chapter)   # 期间对照
    tm    = analyze_table(table_ref, ctx)               # 单表语义
    tms   = analyze_document(doc_model)                 # 全文档
"""

import datetime
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional

import docx_ir

# ---------------------------------------------------------------- 词表

# 列角色：按优先级匹配
COLUMN_ROLE_RULES = [
    ('价值', ['账面价值', '净额', '净值', '账面净值']),
    ('准备', ['坏账准备', '减值准备', '跌价准备', '信用减值准备', '准备']),
    ('折旧', ['累计折旧', '累计摊销', '折旧', '摊销']),
    ('原值', ['账面原值', '原值', '原价']),
    ('余额', ['账面余额', '余额']),
    ('比例', ['计提比例', '坏账率', '比例', '占比']),
    ('增加', ['本期增加', '增加']),
    ('减少', ['本期减少', '减少']),
    ('金额', ['变动金额', '金额', '成本', '收入']),
]

# "金额"这类过于泛化的词只作兜底：它上面一级往往才说明是什么金额
GENERIC_ROLE_WORDS = {'金额', '数额', '数'}
RATIO_ROLES = {'比例'}


def _match_role(text):
    """对单个名称做角色匹配，返回 (角色, 是否泛化词)。"""
    if not text:
        return None, False
    for role, words in COLUMN_ROLE_RULES:
        for w in words:
            if w in text:
                return role, w in GENERIC_ROLE_WORDS
    return None, False


def classify_role(path):
    """把表头路径归到一个角色。

    从**最末一级**往上找：最末一级最具体。
    例如路径 ['期末余额','账面余额','金额']：
        末级「金额」太泛 → 往上一层「账面余额」→ 余额
    再如 ['期末余额','账面余额','比例']：
        末级「比例」→ 直接判为比例（不会被上层的"账面余额"抢走）
    """
    for name in reversed([p for p in path if p]):
        role, generic = _match_role(name)
        if role is None:
            continue
        if generic:
            continue          # 泛化词跳过，继续往上层找更强的
        return role
    # 全是泛化词或没匹配到 → 看看有没有"金额"可用
    for name in reversed([p for p in path if p]):
        role, _ = _match_role(name)
        if role:
            return role
    return None

# 期间关键词 → 语义键（各家模板用词不同：期末/年末、期初/年初、本期/本年…）
PERIOD_KEYS = {
    '本期末': '期末', '期末余额': '期末', '期末': '期末',
    '本年末': '期末', '年末余额': '期末', '年末': '期末',
    '期初余额': '期初', '期初': '期初', '年初余额': '期初', '年初': '期初',
    '上期末': '期初', '上年末': '期初',
    '本期': '本期', '本年': '本期', '本期发生额': '本期',
    '上期': '上期', '上年': '上期',
}

RE_DATE = re.compile(r'(\d{4})\s*[年\-/.]\s*(\d{1,2})\s*[月\-/.]\s*(\d{1,2})')
# 「2026年1月-9月」这种区间：取期末日（9月末）
RE_DATE_RANGE = re.compile(
    r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*[-—~～至到]\s*(\d{1,2})\s*月')
# 「2025年度」「2025年度 收入」：取年末（用 search，不要求整串）
RE_YEAR_ONLY = re.compile(r'(\d{4})\s*年度?(?!\d)')
# 「"期末"指2025年12月31日」
_PERIOD_WORDS = '年末余额|年末|年初余额|年初|期末余额|期末|期初余额|期初|本期|上期'
RE_PERIOD_DEF = re.compile(
    r'[“"\']?(' + _PERIOD_WORDS + r')[”"\']?\s*'
    r'(?:指|为|是|：|:)\s*'
    r'(\d{4})\s*[年\-/.]\s*(\d{1,2})\s*[月\-/.]\s*(\d{1,2})'
)

# 行类型
TOTAL_WORDS = ('合计', '小计', '总计', '总　计', '合　计')
GROUP_WORDS = ('其中：', '其中:', '其中')
SECTION_RE = re.compile(r'^[①②③④⑤⑥⑦⑧⑨⑩]')

NIL_MARKS = {'-', '--', '—', '——', '–', '/', '无', '不适用'}


# ---------------------------------------------------------------- 数据模型

@dataclass
class Column:
    index: int                     # 0 基网格列号
    path: List[str]                # 表头路径，如 ['期末余额','账面余额','金额']
    header_text: str               # 路径拼接后的文字
    role: Optional[str] = None     # 余额/准备/价值/折旧/原值/比例/增加/减少/金额
    period_key: Optional[str] = None   # 期末/期初/本期/上期
    period: Optional[datetime.date] = None
    kind: str = 'amount'           # amount / ratio / text

    @property
    def is_data(self):
        """是否数据列。第 1 列(A列)是行标签，永远不算数据列。"""
        return bool(self.path) and self.index >= 1


@dataclass
class RowInfo:
    row: int                       # 1 基网格行号（表内）
    label: str                     # A 列文字
    kind: str = 'detail'           # header/title/detail/total/group/section/blank
    numeric_count: int = 0


@dataclass
class TableModel:
    chapter: str
    sub_item: str
    ref: object                    # docx_ir / notes_parser.TableRef
    header_rows: int = 0
    title_rows: List[int] = field(default_factory=list)
    columns: List[Column] = field(default_factory=list)
    rows: List[RowInfo] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    @property
    def data_columns(self):
        return [c for c in self.columns if c.is_data]

    @property
    def total_rows(self):
        return [r for r in self.rows if r.kind == 'total']

    def describe(self):
        head = (f'章<{self.chapter}> {self.sub_item}  '
                f'表#{self.ref.idx}  {self.ref.rows}行×{self.ref.cols}列  '
                f'表头{self.header_rows}行  合计行{len(self.total_rows)}')
        lines = [head]
        for c in self.data_columns:
            bits = [f'列{c.index + 1}']
            bits.append(' → '.join(c.path) if c.path else '(无表头)')
            if c.role:
                bits.append(f'[{c.role}]')
            if c.period:
                bits.append(f'@{c.period}')
            elif c.period_key:
                bits.append(f'@{c.period_key}(未定位日期)')
            if c.kind == 'ratio':
                bits.append('(比例)')
            lines.append('   ' + '  '.join(bits))
        return '\n'.join(lines)


# ---------------------------------------------------------------- 工具

def is_numeric_text(text):
    v = docx_ir.parse_cell_value(text)
    return isinstance(v, (int, float))


def _cell_text(grid, r, c):
    if r >= len(grid) or c >= len(grid[r]):
        return '', '?'
    cell = grid[r][c]
    return cell.get('t', ''), cell.get('m', ' ')


def _row_stats(grid, r):
    """返回 (非空格数, 数字格数, 空值标记数)。"""
    non_empty = numeric = nil = 0
    for cell in grid[r]:
        if cell['m'] != ' ':
            continue
        text = (cell['t'] or '').strip()
        if not text:
            continue
        non_empty += 1
        if text in NIL_MARKS:
            nil += 1
            continue
        if is_numeric_text(text):
            numeric += 1
    return non_empty, numeric, nil


def _filled_row(grid, r):
    """把一行按横向合并展开：'h'/'x' 继承左侧文字。返回 list[str]。"""
    out = []
    last = ''
    for cell in grid[r]:
        mark = cell['m']
        text = cell['t']
        if mark in ('h', 'x'):
            out.append(last)          # 横向合并：继承左边
            continue
        if mark in ('v', '?'):
            out.append('')            # 纵向合并：本行没有新信息
            last = text or last
            continue
        out.append(text)
        if text:
            last = text
    return out


# ---------------------------------------------------------------- 期间

def build_chapter_context(doc_model, chapter):
    """从章的正文里读出「期末/期初/本期」对应的具体日期。"""
    ctx = {'period_defs': {}, 'warnings': []}
    texts = []
    for sub in chapter.sub_items:
        texts.extend(sub.paragraphs)
    # 章标题下、第一个子项目前的说明段也要看
    for pos in range(chapter.start, chapter.end):
        it = doc_model.stream[pos]
        if it['kind'] == 'p' and it['text']:
            texts.append(it['text'])

    for t in texts:
        for m in RE_PERIOD_DEF.finditer(t or ''):
            key = PERIOD_KEYS.get(m.group(1))
            if not key:
                continue
            try:
                d = datetime.date(int(m.group(2)), int(m.group(3)), int(m.group(4)))
            except ValueError:
                continue
            ctx['period_defs'].setdefault(key, d)
    if not ctx['period_defs']:
        ctx['warnings'].append('未找到「期末/期初」与具体日期的对照句')
    return ctx


def resolve_period(header_text, ctx):
    """从表头文字解析期间，返回 (period_key, date)。

    区间（2026年1-9月）必须**先于**日期判断：否则日期正则会把它吞成
    "2026年1月9日"（2026-01-09）。
    """
    text = header_text or ''
    m = RE_DATE_RANGE.search(text)
    if m:
        y, m2 = int(m.group(1)), int(m.group(3))
        if 1 <= m2 <= 12:
            day = 31 if m2 in (1, 3, 5, 7, 8, 10, 12) else \
                30 if m2 in (4, 6, 9, 11) else 28
            return None, datetime.date(y, m2, day)
    m = RE_DATE.search(text)
    if m:
        try:
            return None, datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass
    # 「2025年度」→ 年末
    m = RE_YEAR_ONLY.search(text.strip())
    if m:
        return None, datetime.date(int(m.group(1)), 12, 31)
    for word, key in PERIOD_KEYS.items():
        if word in text:
            return key, ctx['period_defs'].get(key)
    return None, None


# ---------------------------------------------------------------- 表头

def detect_header_rows(grid):
    """返回 (title_rows, header_rows)。

    规则：
      · 开头只有 1 个非空格、且没有数字的行 → 视为表标题，跳过
      · 之后「非空格 >= 2、没有数字、也没有空值标记(-)」的连续行 → 表头
        （数据行全是「-」的情况很常见，不能因为没数字就当表头）
      · 表头最多 4 行
    """
    titles = []
    i = 0
    n = len(grid)
    while i < n:
        non_empty, numeric, nil = _row_stats(grid, i)
        if non_empty <= 1 and numeric == 0:
            txt = next((c['t'] for c in grid[i] if c['m'] == ' ' and c['t']), '')
            if txt:
                titles.append(i + 1)   # 1 基行号
            i += 1
            continue
        break

    header = 0
    while i < n and header < 4:
        non_empty, numeric, nil = _row_stats(grid, i)
        if non_empty >= 2 and numeric == 0 and nil == 0:
            header += 1
            i += 1
            continue
        break
    return titles, header


def build_columns(grid, header_rows):
    """按表头行拼出每列的路径。"""
    if header_rows <= 0:
        return []
    filled = [_filled_row(grid, r) for r in range(header_rows)]
    n_cols = max((len(x) for x in filled), default=0)

    columns = []
    for c in range(n_cols):
        path = []
        for row in filled:
            val = row[c] if c < len(row) else ''
            val = (val or '').strip()
            if val and (not path or path[-1] != val):
                path.append(val)
        columns.append(Column(
            index=c,
            path=path,
            header_text=' '.join(path),
        ))
    return columns


# ---------------------------------------------------------------- 行

def classify_rows(grid, header_rows):
    rows = []
    for r in range(len(grid)):
        text, _ = _cell_text(grid, r, 0)
        label = (text or '').strip()
        non_empty, numeric, nil = _row_stats(grid, r)

        if r < header_rows:
            kind = 'header'
        elif not label and non_empty == 0:
            kind = 'blank'
        elif any(label.startswith(w) or label == w for w in TOTAL_WORDS):
            kind = 'total'
        elif label.endswith('小计') or label.endswith('合计'):
            # 「主营业务小计」这类分块小计：它自己已代表下面的明细，
            # 上级合计只取它，不能把它的明细再加一遍
            kind = 'subtotal'
        elif label.startswith(GROUP_WORDS):
            kind = 'group'
        elif SECTION_RE.match(label):
            kind = 'section'
        elif not label:
            kind = 'blank'
        else:
            kind = 'detail'

        rows.append(RowInfo(row=r + 1, label=label, kind=kind, numeric_count=numeric))
    return rows


# ---------------------------------------------------------------- 单表

def analyze_table(table_ref, ctx, chapter_name='', sub_item=''):
    grid = table_ref.grid
    title_rows, header_rows = detect_header_rows(grid)
    columns = build_columns(grid, header_rows)
    rows = classify_rows(grid, header_rows)

    # 表标题行标出来
    title_set = set(title_rows)
    for r in rows:
        if r.row in title_set:
            r.kind = 'title'

    warnings = []
    if header_rows == 0:
        warnings.append('没识别出表头（前几行都不像列名）')

    # 哪些列里真的有数字（用于判断没角色的列是金额列还是文字列）
    data_rows = [r.row - 1 for r in rows
                 if r.kind in ('detail', 'total', 'subtotal', 'group')]

    def _has_values(col_index):
        for r in data_rows:
            if col_index >= len(grid[r]):
                continue
            cell = grid[r][col_index]
            if cell['m'] != ' ':
                continue
            if isinstance(docx_ir.parse_cell_value(cell['t']), (int, float)):
                return True
        return False

    numeric_cols = {c.index for c in columns if _has_values(c.index)}

    for c in columns:
        c.role = classify_role(c.path) if c.is_data else None
        if c.is_data:
            c.period_key, c.period = resolve_period(c.header_text, ctx)
        if c.role in RATIO_ROLES:
            c.kind = 'ratio'
        elif not c.path:
            c.kind = 'text'
        elif c.index in numeric_cols:
            c.kind = 'amount'
        else:
            c.kind = 'text'          # 既没角色又没数字 → 文字列（如"逾期时间"）

    # 修正合并表头与数据列错位：某列**有数据却没有期间**，而旁边一列
    # **有期间却没有数据**（Word 合并单元格跨列与数据列不一致时会出现），
    # 就把期间挪到有数据的那一列。必须在期间解析之后做。
    for c in columns:
        if not c.is_data or c.index not in numeric_cols:
            continue
        if c.period is not None or c.period_key is not None:
            continue
        for nb in (c.index - 1, c.index + 1):
            donor = next((x for x in columns if x.index == nb), None)
            if donor is None or not donor.is_data:
                continue
            if donor.index in numeric_cols:
                continue          # 接手的对象必须是没有数据的空列
            if donor.period is None and donor.period_key is None:
                continue
            c.period, c.period_key = donor.period, donor.period_key
            if c.role is None:
                c.role = donor.role
            donor.period = donor.period_key = None
            break

    for c in columns:
        if c.is_data and c.role is not None and c.period is None and c.period_key is None:
            # 金额/比例列但表头没写期间：只作提示，不算"结构没认出来"
            warnings.append(f'（期间待确认）列{c.index + 1}「{c.header_text or "空"}」')

    return TableModel(
        chapter=chapter_name,
        sub_item=sub_item,
        ref=table_ref,
        header_rows=header_rows,
        title_rows=title_rows,
        columns=columns,
        rows=rows,
        warnings=warnings,
    )


def analyze_document(doc_model, hint=None, stats=None):
    """对整份附注的所有表做语义解析。

    hint: 可选的"模板提示"回调
          hint(章名, 子项目名, 该子项目下第几张表, 数据列数) -> [角色或 None, ...]
          只在自主识别没判出角色时用来兜底——模板只是加速，不是必需。
    stats: 可选 dict，用于统计模板补了多少列。
    """
    out = []
    for ch in doc_model.chapters:
        ctx = build_chapter_context(doc_model, ch)
        for sub in ch.sub_items:
            for n, t in enumerate(sub.tables, 1):
                tm = analyze_table(t, ctx, ch.name, sub.name)
                if hint is not None:
                    _apply_template_hint(tm, hint, ch.name, sub.name, n, stats)
                out.append(tm)
    return out


def _apply_template_hint(tm, hint, chapter_name, sub_name, table_seq, stats=None):
    """用模板骨架给没判出角色的列补角色。只补空缺，不覆盖已有判断。"""
    data_cols = [c for c in tm.columns if c.is_data]
    if not data_cols:
        return
    try:
        roles = hint(chapter_name, sub_name, table_seq, len(data_cols))
    except Exception:
        return
    if not roles:
        return
    for col, role in zip(data_cols, roles):
        if not role or col.role:
            continue
        col.role = role
        if role in RATIO_ROLES:
            col.kind = 'ratio'
        elif col.kind == 'text':
            col.kind = 'amount'
        if stats is not None:
            stats['template_roles'] = stats.get('template_roles', 0) + 1


# ---------------------------------------------------------------- CLI

def _main():
    import sys
    import notes_parser

    path = sys.argv[1] if len(sys.argv) > 1 else None
    if not path:
        print(__doc__)
        return 2

    model = notes_parser.parse_notes(path, work_dir='_work/parse')
    for ch in model.chapters:
        ctx = build_chapter_context(model, ch)
        print(f'章<{ch.name}> 期间对照: {ctx["period_defs"]}  {ctx["warnings"]}')

    tables = analyze_document(model)
    print(f'\n共解析 {len(tables)} 张表')

    stats = {'有表头': 0, '有合计行': 0, '有角色列': 0}
    for tm in tables:
        if tm.header_rows:
            stats['有表头'] += 1
        if tm.total_rows:
            stats['有合计行'] += 1
        if any(c.role for c in tm.columns):
            stats['有角色列'] += 1
    print(f'统计: {stats}')

    for tm in tables[:4]:
        print()
        print(tm.describe())
    return 0


if __name__ == '__main__':
    raise SystemExit(_main())
