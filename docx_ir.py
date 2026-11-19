# -*- coding: utf-8 -*-
"""
docx_ir.py — 附注文档的「结构中间层」（IR）。

设计目标
--------
1. **不依赖案例**：不写死任何科目、行号、样式 ID。
2. **快**：.doc 先用 Word 转一次 .docx，之后全部用 python-docx 离线读。
   （逐单元格走 Word COM 一份 2MB 模板要 29 分钟；转 docx 后秒级）
3. **保真**：段落与表格按文档物理顺序排；表格保留完整网格坐标与合并标记。

已用 6 份文档（5 份模板 + 1 份真实附注）验证：
    「章」按关键词 + 大纲级别锚定；
    「子项目级别」= 该章区间内最小的、大于章级别的大纲级别；
    该规则在 6 份文档上 100% 正确（含模板无编号、真实附注用 H4 的情形）。

对外主要接口
------------
    load_stream(docx_path)                  -> list[dict]      文档流
    ensure_docx(src_path, work_dir=None)    -> str             .doc/.docx -> .docx
    find_note_chapters(stream)              -> list[Chapter]   附注章
    find_sub_items(stream, chapter)         -> list[SubItem]   子项目
    iter_chapter_rows(stream, chapter)      -> list[list]      按行铺平的表格数据
    parse_cell_value(text)                  -> None|float|date|str
"""

import os
import re
import shutil
import tempfile

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

import docx_numbering

# ---------------------------------------------------------------- 关键词

# 附注章标题关键词（宽一点：模板里叫法各不相同）
NOTE_CHAPTER_KEYS = (
    '主要项目附注', '主要项目注释', '财务报表项目注释', '财务报表项目附注',
    '报表项目附注', '报表项目注释', '报表主要项目附注', '主要财务报表项目附注',
)
# 命中这些词的章不是主要项目附注
NOTE_CHAPTER_EXCLUDE = ('会计政策', '变更', '差错', '编制基础', '补充资料')

# 合并/母公司范围词
SCOPE_PARENT_KEYS = ('母公司', '公司财务报表')
SCOPE_CONSOL_KEYS = ('合并',)

RE_CN_NUMBER = re.compile(r'^\s*[一二三四五六七八九十百]+\s*[、．.]')
RE_ARAB_NUMBER = re.compile(r'^\s*(\d+)\s*[、．.]\s*(\S.*)$')
RE_PAREN_NUMBER = re.compile(r'^\s*[（(]\s*(\d+)\s*[)）]\s*(\S.*)$')

# 数值/日期
RE_NUM_COMMA = re.compile(r'^[-+]?[\d,]+(\.\d+)?$')
RE_PERCENT = re.compile(r'^[-+]?[\d,]*\.?\d+%$')
RE_DATE_CN = re.compile(r'^(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日$')
RE_DATE_SLASH = re.compile(r'^(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})$')
NIL_MARKS = {'-', '--', '—', '——', '–', '/'}


# ---------------------------------------------------------------- 读文档

def _qc(tc):
    """取表格单元格里的纯文本（不含段落标记）。"""
    return ''.join(t.text or '' for t in tc.iter(qn('w:t'))).strip()


def _outline_of_ppr(pPr):
    if pPr is None:
        return None
    el = pPr.find(qn('w:outlineLvl'))
    if el is None:
        return None
    try:
        return int(el.get(qn('w:val')))
    except (TypeError, ValueError):
        return None


def _effective_outline(paragraph):
    """段落自身 → 样式链，取第一个 outlineLvl（0 基）。无则 None。"""
    lvl = _outline_of_ppr(paragraph._p.pPr)
    if lvl is not None:
        return lvl
    style = paragraph.style
    for _ in range(10):
        if style is None:
            break
        lvl = _outline_of_ppr(style.element.find(qn('w:pPr')))
        if lvl is not None:
            return lvl
        style = style.base_style
    return None


def _dump_table(tbl, tidx):
    """表格 → 网格。merge 标记：' ' 正常 / 'h' 左并 / 'v' 上并 / 'x' 双向。

    只保留「合并区左上角」的文本，被覆盖的位置留空，避免同一内容重复出现。

    实现要点：**不要**用 python-docx 的 `table.cell(r, c)` 再比较 `id(tc)` 判断合并。
    lxml 的元素代理会被回收、`id()` 会被复用，导致正常单元格被误判成"被合并覆盖"，
    金额列会被整列吃掉。这里直接读 `<w:tbl>` 的 XML 真值：
        <w:tr> -> <w:tc>，每格看 <w:tcPr><w:gridSpan>（横向跨列）
                                    <w:vMerge>（纵向合并，restart=起点，缺省/continue=被覆盖）
        行首跨列看 <w:trPr><w:gridBefore>
    """
    tbl_el = tbl._tbl

    grid_el = tbl_el.find(qn('w:tblGrid'))
    col_el = grid_el.findall(qn('w:gridCol')) if grid_el is not None else []
    n_cols = len(col_el)
    widths = []
    for c in col_el:
        try:
            widths.append(int(c.get(qn('w:w'), '0')))
        except (TypeError, ValueError):
            widths.append(0)
    # Word 里常有 100 多 twips 的"分隔列"，它会把合并格和数据列错开一格。
    # 只丢**又窄又完全空白**的列：先把所有格子的文字落在哪些原始列上扫一遍。
    tr_list_pre = tbl_el.findall(qn('w:tr'))
    content_cols = set()
    for _tr in tr_list_pre:
        _col = 0
        _trPr = _tr.find(qn('w:trPr'))
        if _trPr is not None:
            _gb = _trPr.find(qn('w:gridBefore'))
            if _gb is not None:
                try:
                    _col = int(_gb.get(qn('w:val'), '0'))
                except (TypeError, ValueError):
                    _col = 0
        for _tc in _tr.findall(qn('w:tc')):
            _span = 1
            _tcPr = _tc.find(qn('w:tcPr'))
            if _tcPr is not None:
                _gs = _tcPr.find(qn('w:gridSpan'))
                if _gs is not None:
                    try:
                        _span = max(1, int(_gs.get(qn('w:val'), '1')))
                    except (TypeError, ValueError):
                        _span = 1
            _txt = _qc(_tc)
            if _txt.strip():
                for _k in range(_span):
                    content_cols.add(_col + _k)
            _col += _span

    narrow = set()
    if widths:
        pos = sorted(w for w in widths if w > 0)
        median = pos[len(pos) // 2] if pos else 0
        if median > 0:
            narrow = {i for i, w in enumerate(widths)
                      if 0 < w < median * 0.25 and w < 250}

    def _new_index(old_i):
        if not narrow:
            return old_i
        return sum(1 for i in range(old_i) if i not in narrow)

    tr_list = tbl_el.findall(qn('w:tr'))
    n_rows = len(tr_list)

    grid = []
    for tr in tr_list:
        # 行首因纵向合并而空出的列数
        grid_before = 0
        trPr = tr.find(qn('w:trPr'))
        if trPr is not None:
            gb = trPr.find(qn('w:gridBefore'))
            if gb is not None:
                try:
                    grid_before = int(gb.get(qn('w:val'), '0'))
                except (TypeError, ValueError):
                    grid_before = 0

        row = [{'t': '', 'm': '?'} for _ in range(grid_before)]
        col = grid_before
        for tc in tr.findall(qn('w:tc')):
            span = 1
            vmerge = None
            tcPr = tc.find(qn('w:tcPr'))
            if tcPr is not None:
                gs = tcPr.find(qn('w:gridSpan'))
                if gs is not None:
                    try:
                        span = max(1, int(gs.get(qn('w:val'), '1')))
                    except (TypeError, ValueError):
                        span = 1
                vm = tcPr.find(qn('w:vMerge'))
                if vm is not None:
                    vmerge = vm.get(qn('w:val')) or 'continue'

            is_cont = vmerge == 'continue'
            text = '' if is_cont else _qc(tc)

            # 把真实的合并跨度记下来：后面还原表头必须靠它，
            # 只靠 'h' 标记按位置推断会错位（Word 里合并格起点常偏一格）
            # 这一格覆盖的原始列号
            cover = [col + k for k in range(span)]
            keep = [i for i in cover if i not in narrow]
            if not keep:
                col += span
                continue      # 整格都落在窄分隔列里，直接丢掉，不占列
            row.append({'t': text, 'm': 'v' if is_cont else ' ', 'span': len(keep)})
            for _k in range(1, len(keep)):
                row.append({'t': '', 'm': 'x' if is_cont else 'h', 'span': 0})
            col += span

        if n_cols < len(row):
            n_cols = len(row)
        grid.append(row)

    # 统一每行长度
    for row in grid:
        while len(row) < n_cols:
            row.append({'t': '', 'm': '?'})

    return {'idx': tidx, 'rows': n_rows, 'cols': n_cols, 'grid': grid}


def numbers_sidecar(docx_path):
    """渲染编号的旁挂文件路径（与 docx 同目录同名）。"""
    return os.path.splitext(os.path.abspath(docx_path))[0] + '.numbers.json'


def dump_rendered_numbers(word_doc, docx_path):
    """让 Word 报出每个段落的**实际显示编号**，写成 json 旁挂文件。

    word_doc 是已经打开的 Word.Document（COM 对象）。
    ListFormat.ListString 就是用户在 Word 里看到的"五、""1、"。
    """
    import json
    mapping = {}
    try:
        paragraphs = word_doc.Paragraphs
        total = paragraphs.Count
    except Exception:
        return None

    for i in range(1, total + 1):
        try:
            p = paragraphs(i)
            text = (p.Range.Text or '').replace('\r', '').replace('\x07', '').strip()
            if not text:
                continue
            num = (p.Range.ListFormat.ListString or '').strip()
            if num:
                mapping[text] = num
        except Exception:
            continue

    try:
        with open(numbers_sidecar(docx_path), 'w', encoding='utf-8') as f:
            json.dump(mapping, f, ensure_ascii=False, indent=1)
    except Exception:
        pass
    return mapping


def load_rendered_numbers(docx_path):
    """读回 Word 报的编号映射（没有就返回空 dict）。"""
    import json
    path = numbers_sidecar(docx_path)
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def load_stream(docx_path):
    """按文档物理顺序返回 段落 / 表格 流。

    段落额外带 auto_number / auto_number_value / num_ilvl：
    Word 自动编号不在正文文本里，需要从 numbering.xml 还原（见 docx_numbering）。
    """
    doc = Document(docx_path)

    # 一次性按顺序算好自动编号（模板的章节号全在这里面）
    try:
        number_map = docx_numbering.walk_labels(doc)
    except Exception:
        number_map = {}

    stream = []
    t_idx = 0
    p_idx = 0
    rendered = load_rendered_numbers(docx_path)     # Word 报的"实际显示编号"
    for child in doc.element.body.iterchildren():
        if child.tag == qn('w:p'):
            para = Paragraph(child, doc)
            try:
                style_id = para.style.style_id or ''
            except Exception:
                style_id = ''
            num = number_map.get(child)
            text = para.text.strip()
            label = num['label'] if num else None
            # 以 Word 实际渲染的编号为准（逐字匹配段落文字）
            if rendered and label and text in rendered:
                label = rendered[text]
            stream.append({
                'kind': 'p',
                'i': p_idx,
                'text': text,
                'style_id': style_id,
                'outline': _effective_outline(para),
                'auto_number': label,
                'auto_number_value': num['value'] if num else None,
                'num_ilvl': num['ilvl'] if num else None,
            })
            p_idx += 1
        elif child.tag == qn('w:tbl'):
            stream.append({'kind': 'table', **_dump_table(Table(child, doc), t_idx)})
            t_idx += 1
    return stream


# ---------------------------------------------------------------- 转 docx

def _run_com_with_timeout(fn, timeout, what):
    """在独立 STA 线程里跑 Office 自动化，超时就报可操作的错误，避免界面卡死。"""
    import threading
    import time

    box = {}

    def target():
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass
        try:
            box['value'] = fn()
        except Exception as exc:            # noqa: BLE001
            box['error'] = exc

    th = threading.Thread(target=target, daemon=True)
    th.start()
    deadline = time.time() + timeout
    while time.time() < deadline and th.is_alive():
        th.join(0.3)

    if th.is_alive():
        raise RuntimeError(
            f'{what}超过 {timeout} 秒没有响应。\n'
            f'请先关闭所有 Word / Excel 窗口后重试；'
            f'若仍不行，把附注另存为 .docx、TB 另存为 .xlsx 再放进来。')
    if 'error' in box:
        exc = box['error']
        raise RuntimeError(
            f'{what}失败：{exc}\n'
            f'请确认已安装 Office；或先手动把文件另存为 .docx / .xlsx。')
    return box.get('value')


def ensure_docx(src_path, work_dir=None):
    """把 .doc/.docx 归一成 .docx 路径。.doc 走 Word COM 转一次。"""
    src_path = os.path.abspath(src_path)
    if src_path.lower().endswith('.docx'):
        return src_path
    if not src_path.lower().endswith('.doc'):
        raise ValueError(f'不支持的附注格式: {src_path}')

    work_dir = os.path.abspath(work_dir) if work_dir else tempfile.mkdtemp(prefix='notes_docx_')
    os.makedirs(work_dir, exist_ok=True)
    dst = os.path.join(work_dir, os.path.splitext(os.path.basename(src_path))[0] + '.docx')
    if os.path.exists(dst) and os.path.getmtime(dst) >= os.path.getmtime(src_path):
        return dst

    import win32com.client as win32
    import com_env

    # 中文路径在打包环境里交给 Word 会出问题，先摆到纯英文暂存路径
    src_use, staged = com_env.ascii_staging(src_path, '.doc', work_dir)
    early = com_env.staging_dir(work_dir)
    if early:
        import hashlib
        tag = hashlib.md5(os.path.abspath(src_path).encode('utf-8')).hexdigest()[:10]
        out_stage = os.path.join(early, f'out_{tag}.docx')
    else:
        out_stage = dst

    def _convert():
        word = None
        with com_env.sanitized():
            word = win32.DispatchEx('Word.Application')
        try:
            word.Visible = False
            word.DisplayAlerts = 0
            try:
                word.AutomationSecurity = 3
            except Exception:
                pass
            # 注意：Word 的当前工作目录不是本进程的当前目录，
            # Open/SaveAs2 必须传绝对路径，否则报"这不是有效文件名"。
            doc = word.Documents.Open(os.path.abspath(src_use), ConfirmConversions=False,
                                      ReadOnly=False, AddToRecentFiles=False)
            # 让 Word 把自动编号"烤"成正文文字：这样 docx 里的编号就是
            # 用户在 Word 里看到的那一个（numbering.xml 的起始值不一定等于渲染结果）。
            # 顺带把编号问题彻底消灭——后面读到的都是字面编号。
            try:
                doc.ConvertNumbersToText()
            except Exception:
                pass
            doc.SaveAs2(os.path.abspath(out_stage), FileFormat=16)
            doc.Close(False)
        finally:
            try:
                word.Quit()
            except Exception:
                pass
        return out_stage

    _run_com_with_timeout(_convert, 300, 'Word 转换附注')
    if os.path.abspath(out_stage) != os.path.abspath(dst):
        import shutil
        shutil.copy2(out_stage, dst)
    return dst


# ---------------------------------------------------------------- 章节结构

def is_note_chapter_text(text):
    if not text or not any(k in text for k in NOTE_CHAPTER_KEYS):
        return False
    return not any(x in text for x in NOTE_CHAPTER_EXCLUDE)


def chapter_scope(text):
    """判断这一章是 合并 还是 母公司。"""
    if any(k in text for k in SCOPE_PARENT_KEYS):
        return '母公司'
    if any(k in text for k in SCOPE_CONSOL_KEYS):
        return '合并'
    return '单体'


def find_note_chapters(stream):
    """找出所有「主要项目附注」章，返回 [{name, outline, scope, start, end}]。"""
    headings = [(n, it) for n, it in enumerate(stream)
                if it['kind'] == 'p' and it['text'] and it.get('outline') is not None]

    chapters = []
    for pos, (n, it) in enumerate(headings):
        if not is_note_chapter_text(it['text']):
            continue
        level = it['outline']
        end = len(stream)
        for n2, it2 in headings[pos + 1:]:
            if it2['outline'] <= level:
                end = n2
                break
        chapters.append({
            'name': it['text'],
            'outline': level,
            'scope': chapter_scope(it['text']),
            'start': n,
            'end': end,
        })
    return chapters


def sub_item_level(stream, chapter):
    """子项目级别 = 章区间内最小的、大于章级别的大纲级别。"""
    levels = [it['outline'] for it in stream[chapter['start'] + 1:chapter['end']]
              if it['kind'] == 'p' and it.get('outline') is not None
              and it['outline'] > chapter['outline']]
    return min(levels) if levels else None


def find_sub_items(stream, chapter):
    """返回 [{idx, name, seq, number, start, end}]。

    number: 标题里写的编号（如 "1、货币资金" -> 1）；模板里没有编号则为 None。
    seq:    在章内的出现顺序（1 基）。写进成果表时用 seq，避免被原稿编号带偏。
    """
    level = sub_item_level(stream, chapter)
    if level is None:
        return []

    heads = [(n, it) for n, it in enumerate(stream)
             if chapter['start'] < n < chapter['end']
             and it['kind'] == 'p' and it['text']
             and it.get('outline') == level]

    items = []
    for k, (n, it) in enumerate(heads):
        end = heads[k + 1][0] if k + 1 < len(heads) else chapter['end']
        m = RE_ARAB_NUMBER.match(it['text'])
        number = int(m.group(1)) if m else None
        name = (m.group(2).strip() if m else it['text'].strip())
        items.append({
            'idx': n,
            'name': name,
            'raw_title': it['text'],
            'number': number,
            'seq': k + 1,
            'start': n + 1,
            'end': end,
        })
    return items


# ---------------------------------------------------------------- 值解析

RE_TRAIL_MINUS = re.compile(r'^([\d,]+(?:\.\d+)?)\s*-\s*$')


def is_trailing_minus(text):
    """判断是不是「9,900.00-」这种尾随负号写法。"""
    return bool(RE_TRAIL_MINUS.match(str(text or '').strip()))


def parse_cell_value(text):
    """把表格文字转成 Excel 里能用的值。

    '7,921,661.42' -> 7921661.42
    '100.00%'      -> 1.0
    '2026年9月30日' -> datetime.date(2026, 9, 30)
    '- / --'        -> '-'（保留，表示该格为无，不是 0）
    '9,900.00-'     -> -9900.00（尾随负号，数值照算，另外单独提示删掉 -）
    ''              -> None
    """
    if text is None:
        return None
    s = str(text).strip()
    if not s:
        return None
    if s in NIL_MARKS:
        return '-'

    m = RE_TRAIL_MINUS.match(s)
    if m:
        try:
            return -float(m.group(1).replace(',', ''))
        except ValueError:
            return s

    if RE_PERCENT.match(s):
        try:
            return float(s.rstrip('%').replace(',', '')) / 100.0
        except ValueError:
            return s

    m = RE_DATE_CN.match(s) or RE_DATE_SLASH.match(s)
    if m:
        import datetime
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return s

    if RE_NUM_COMMA.match(s):
        try:
            return float(s.replace(',', ''))
        except ValueError:
            return s
    return s


# ---------------------------------------------------------------- 铺平

def iter_chapter_rows(stream, chapter):
    """把章区间内的段落与表格按顺序铺平成 Excel 行。

    规则（与原始 Word 版式对应）：
      - 一个段落 = 一行，文字放 A 列（空段落也占一行，保持行数与文档一致）
      - 一张表 = 表格行数那么多行，一个单元格一列
      - 合并区只写左上角，其余留空 —— 列位不塌陷，能和原表对上
    """
    rows = []
    for it in stream[chapter['start']:chapter['end']]:
        if it['kind'] == 'p':
            rows.append({'source': f"p{it['i']}", 'cells': [it['text']]})
        else:
            for r_idx, row in enumerate(it['grid'], 1):
                values = []
                for cell in row:
                    values.append(parse_cell_value(cell['t']) if cell['m'] == ' ' else None)
                while values and values[-1] is None:
                    values.pop()
                rows.append({'source': f"T{it['idx']}R{r_idx}", 'cells': values})
    return rows


def copy_docx_into(src, dst_dir):
    os.makedirs(dst_dir, exist_ok=True)
    dst = os.path.join(dst_dir, os.path.basename(src))
    if os.path.abspath(src) != os.path.abspath(dst):
        shutil.copy2(src, dst)
    return dst
