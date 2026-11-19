# -*- coding: utf-8 -*-
"""
inspect_doc.py — 用 Word COM 把 .doc/.docx 的「顶层段落 + 表格」按文档顺序导成 JSON/TXT。

用途：给"附注勾稽引擎"摸清模板与真实附注的结构（章节、小节标题、表格 rows x cols、单元格文本）。

用法：
    python tools/inspect_doc.py "附注模板\\一般企业模板6.4财务报表附注.doc" [输出前缀]

输出（与源文件同目录）：
    <前缀>.structure.json  —— 完整结构，供程序分析
    <前缀>.structure.txt   —— 人读版
"""

import json
import os
import sys
import time
import traceback

import win32com.client as win32

WD_ALERTS_NONE = 0
WD_AUTOMATION_SECURITY_DISABLE = 3

# 表格单元格分隔/行结束符
CELL_MARK = '\x07'
PARA_MARK = '\r'


def _clean(s):
    if s is None:
        return ''
    return s.replace(CELL_MARK, '').replace(PARA_MARK, '').replace('\x0b', '\n').strip()


def open_word():
    word = win32.DispatchEx('Word.Application')
    word.Visible = False
    word.DisplayAlerts = WD_ALERTS_NONE
    try:
        word.AutomationSecurity = WD_AUTOMATION_SECURITY_DISABLE
    except Exception:
        pass
    return word


def dump_doc(path, log):
    word = open_word()
    doc = None
    started = time.time()
    try:
        doc = word.Documents.Open(
            os.path.abspath(path),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        log(f'opened in {time.time() - started:.1f}s: paragraphs={doc.Paragraphs.Count} '
            f'tables={doc.Tables.Count} sections={doc.Sections.Count}')

        # 1) 顶层段落（不在任何表格内的段落），带样式名，用于识别大标题/小节标题
        para_list = []
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            try:
                in_table = p.Range.Tables.Count > 0
            except Exception:
                in_table = False
            if in_table:
                continue
            para_list.append({
                'start': p.Range.Start,
                'style': str(p.Style.NameLocal),
                'outline': int(p.OutlineLevel),
                'text': _clean(p.Range.Text),
            })
        log(f'top-level paragraphs: {len(para_list)}')

        # 2) 表格：逐单元格取 RowIndex/ColumnIndex/Text（只取真实存在的单元格，
        #    合并区被覆盖的位置不会出现，正好符合"合并不占列位"的需要）
        table_list = []
        for ti in range(1, doc.Tables.Count + 1):
            t = doc.Tables(ti)
            rows = t.Rows.Count
            cols = t.Columns.Count
            cells = []
            try:
                for c in t.Range.Cells:
                    cells.append({
                        'r': int(c.RowIndex),
                        'c': int(c.ColumnIndex),
                        'text': _clean(c.Range.Text),
                    })
            except Exception as exc:
                cells.append({'r': 0, 'c': 0, 'text': f'<CELL READ ERROR: {exc}>'})
            table_list.append({
                'start': t.Range.Start,
                'rows': rows,
                'cols': cols,
                'uniform': bool(t.Uniform),
                'cell_count': len(cells),
                'cells': cells,
            })
            if ti % 20 == 0:
                log(f'  tables {ti}/{doc.Tables.Count} ...')

        doc.Close(False)
        doc = None
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        try:
            word.Quit()
        except Exception:
            pass

    # 3) 合并成按文档顺序排列的流
    stream = []
    for p in para_list:
        stream.append({'kind': 'p', 'start': p['start'], 'style': p['style'],
                       'outline': p['outline'], 'text': p['text']})
    for idx, t in enumerate(table_list):
        stream.append({'kind': 'table', 'start': t['start'], 'table_idx': idx,
                       'rows': t['rows'], 'cols': t['cols'], 'uniform': t['uniform'],
                       'cells': t['cells']})
    stream.sort(key=lambda x: x['start'])

    return {
        'source': os.path.abspath(path),
        'paragraph_count': len(para_list),
        'table_count': len(table_list),
        'stream': stream,
    }


def write_readable(result, out_txt):
    lines = []
    lines.append(f"SOURCE: {result['source']}")
    lines.append(f"top-level paragraphs={result['paragraph_count']}  tables={result['table_count']}")
    lines.append('=' * 100)
    for item in result['stream']:
        if item['kind'] == 'p':
            if not item['text']:
                continue
            ol = item['outline']
            mark = f'[H{ol}]' if ol and ol < 10 else '     '
            lines.append(f"{mark} P  {item['text']}")
        else:
            lines.append(f"--- TABLE#{item['table_idx']}  {item['rows']}行 x {item['cols']}列 "
                         f"uniform={item['uniform']} cells={len(item['cells'])}")
            by_row = {}
            for c in item['cells']:
                by_row.setdefault(c['r'], []).append(c)
            for r in sorted(by_row):
                cells = sorted(by_row[r], key=lambda x: x['c'])
                rendered = ' | '.join(f"{c['c']}:{c['text']}" for c in cells)
                lines.append(f"    R{r}: {rendered}")
    open(out_txt, 'w', encoding='utf-8').write('\n'.join(lines))


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    path = sys.argv[1]
    prefix = sys.argv[2] if len(sys.argv) > 2 else os.path.splitext(path)[0]
    out_json = prefix + '.structure.json'
    out_txt = prefix + '.structure.txt'

    logf = open(prefix + '.structure.log', 'w', encoding='utf-8')

    def log(msg):
        print(msg)
        logf.write(str(msg) + '\n')
        logf.flush()

    try:
        result = dump_doc(path, log)
        json.dump(result, open(out_json, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
        write_readable(result, out_txt)
        log(f'OK -> {out_json}')
        log(f'OK -> {out_txt}')
        return 0
    except Exception:
        log(traceback.format_exc())
        return 1
    finally:
        logf.close()


if __name__ == '__main__':
    sys.exit(main())
