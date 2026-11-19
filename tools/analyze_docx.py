# -*- coding: utf-8 -*-
"""
analyze_docx.py — 离线解析 .docx 的结构（段落流 + 表格网格），输出 JSON/TXT。

为什么不用 Word COM 逐单元格：太慢（2MB 文档 252 张表要 29 分钟）。
.doc 先用 tools/doc_to_docx.py 转成 .docx，这里纯离线读，秒级完成。

关键产出（供通用勾稽引擎使用）：
  - 按文档顺序的 段落 / 表格 流
  - 每段的 样式ID / 样式名 / 有效大纲级别（用于锚定"章"和"子项目"）
  - 每表的 网格行列数 + 每格文本 + 是否被合并覆盖（left/up 同源）

用法：
    python tools/analyze_docx.py <输出前缀> <文件.docx>
    python tools/analyze_docx.py _docx "附注模板"      # 目录下全部 docx
"""

import json
import os
import sys
import time

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph


# ---------------------------------------------------------------- 样式/大纲

def _outline_from_ppr(pPr):
    """从 <w:pPr><w:outlineLvl w:val="n"/> 取大纲级别（0=一级）。"""
    if pPr is None:
        return None
    el = pPr.find(qn('w:outlineLvl'))
    if el is None:
        return None
    val = el.get(qn('w:val'))
    try:
        return int(val)
    except (TypeError, ValueError):
        return None


def effective_outline(paragraph):
    """段落自身 → 样式链 → 取第一个 outlineLvl。返回 0-based，未定义返回 None。"""
    lvl = _outline_from_ppr(paragraph._p.pPr)
    if lvl is not None:
        return lvl
    style = paragraph.style
    seen = 0
    while style is not None and seen < 10:
        sel = style.element
        lvl = _outline_from_ppr(sel.find(qn('w:pPr')))
        if lvl is not None:
            return lvl
        style = style.base_style
        seen += 1
    return None


def cell_text(tc):
    return ''.join(t.text or '' for t in tc.iter(qn('w:t'))).strip()


# ---------------------------------------------------------------- 表格

def dump_table(tbl, tidx):
    """把表格导成网格：rows x cols，每格 text + merge 标记。

    merge 标记：'o' = 合并区左上角(origin)，'h' = 被左侧横向合并覆盖，
               'v' = 被上方纵向合并覆盖，'x' = 两者都是，' ' = 普通单元格。
    """
    rows = len(tbl.rows)
    cols = len(tbl.columns)
    grid = []
    prev_row_ids = [None] * cols
    for r in range(rows):
        row_cells = []
        prev_id = None
        for c in range(cols):
            try:
                cell = tbl.cell(r, c)
                tc = cell._tc
            except (IndexError, ValueError):
                row_cells.append({'t': '', 'm': '?'})
                continue
            tid = id(tc)
            above_same = prev_row_ids[c] == tid
            left_same = prev_id == tid
            if above_same and left_same:
                mark = 'x'
            elif above_same:
                mark = 'v'
            elif left_same:
                mark = 'h'
            else:
                mark = ' '
            # origin 才写文本，覆盖位置留空，避免同一内容重复
            row_cells.append({'t': '' if mark != ' ' else cell_text(tc), 'm': mark})
            prev_id = tid
            prev_row_ids[c] = tid
        grid.append(row_cells)
    return {'idx': tidx, 'rows': rows, 'cols': cols, 'grid': grid}


# ---------------------------------------------------------------- 主流程

def analyze(path):
    doc = Document(path)
    body = doc.element.body
    t_idx = 0
    p_idx = 0
    stream = []

    for child in body.iterchildren():
        if child.tag == qn('w:p'):
            para = Paragraph(child, doc)
            text = para.text.strip()
            style_id = ''
            style_name = ''
            try:
                style_id = para.style.style_id or ''
                style_name = para.style.name or ''
            except Exception:
                pass
            stream.append({
                'kind': 'p',
                'i': p_idx,
                'text': text,
                'style_id': style_id,
                'style_name': style_name,
                'outline': effective_outline(para),
            })
            p_idx += 1
        elif child.tag == qn('w:tbl'):
            tbl = Table(child, doc)
            stream.append({'kind': 'table', **dump_table(tbl, t_idx)})
            t_idx += 1

    return {
        'source': os.path.abspath(path),
        'paragraph_count': p_idx,
        'table_count': t_idx,
        'stream': stream,
    }


def write_readable(res, out_txt):
    lines = [f"SOURCE: {res['source']}",
             f"paragraphs={res['paragraph_count']} tables={res['table_count']}",
             '=' * 100]
    for it in res['stream']:
        if it['kind'] == 'p':
            if not it['text']:
                continue
            ol = it.get('outline')
            mark = f'[H{ol + 1}]' if ol is not None and ol < 9 else '     '
            lines.append(f"{mark} {it['style_id']:<12} {it['text']}")
        else:
            visible = sum(1 for row in it['grid'] for c in row
                          if c['m'] == ' ' and c['t'])
            lines.append(f"--- TABLE#{it['idx']}  {it['rows']}行 x {it['cols']}列 "
                         f"有值格={visible}")
            for r, row in enumerate(it['grid'], 1):
                cells = [f"{c + 1}:{cell['t']}" if cell['m'] == ' ' and cell['t']
                         else f"{c + 1}:·" if cell['m'] == ' ' else None
                         for c, cell in enumerate(row)]
                rendered = ' | '.join(x for x in cells if x)
                lines.append(f"    R{r}: {rendered}")
    open(out_txt, 'w', encoding='utf-8').write('\n'.join(lines))


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        return 2
    prefix = sys.argv[1]
    targets = sys.argv[2:]

    files = []
    for t in targets:
        if os.path.isdir(t):
            for f in sorted(os.listdir(t)):
                if f.lower().endswith('.docx') and not f.startswith('~$'):
                    files.append(os.path.join(t, f))
        else:
            files.append(t)

    for f in files:
        t0 = time.time()
        res = analyze(f)
        base = os.path.splitext(os.path.basename(f))[0]
        out_json = os.path.join(prefix, '_a_' + base + '.json')
        out_txt = os.path.join(prefix, '_a_' + base + '.txt')
        os.makedirs(prefix, exist_ok=True)
        json.dump(res, open(out_json, 'w', encoding='utf-8'), ensure_ascii=False)
        write_readable(res, out_txt)
        print(f'OK {time.time() - t0:.2f}s  {base}  '
              f'paragraphs={res["paragraph_count"]} tables={res["table_count"]}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
