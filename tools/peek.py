# -*- coding: utf-8 -*-
"""
peek.py — 从 analyze_docx 的 JSON 里按区间打印流对象，用来人工核对结构。

用法：
    python tools/peek.py <json> <start> [count]        # 打印流区间
    python tools/peek.py <json> --grep "正则"           # 全文搜索
    python tools/peek.py <json> --numheads <lo> <hi>   # 找"数字编号"标题
    python tools/peek.py <json> --table <idx>          # 打印某张表的网格
    python tools/peek.py <json> --chapters             # 只看章级标题 + 附注章判定
"""

import json
import re
import sys

NOTE_KEYS = ('主要项目附注', '主要项目注释', '财务报表项目注释', '财务报表项目附注',
             '报表项目附注', '报表项目注释')
RE_CN_HEAD = re.compile(r'^\s*[一二三四五六七八九十百]+\s*[、．.]')
RE_ARAB = re.compile(r'^\s*(\d+)\s*[、．.]')


def load(path):
    return json.load(open(path, encoding='utf-8'))


def fmt(it):
    if it['kind'] == 'p':
        ol = it.get('outline')
        flag = '' if ol is None else f'H{ol + 1}'
        return (f"  p{it['i']:<6d} {flag:<4} {it.get('style_id', ''):<8} "
                f"{it['text'][:110]}")
    return f"  >>> TABLE#{it['idx']} {it['rows']}x{it['cols']}"


def is_note_chapter(text):
    return any(k in text for k in NOTE_KEYS)


def main():
    path = sys.argv[1]
    res = load(path)
    stream = res['stream']
    mode = sys.argv[2]

    if mode == '--grep':
        pat = re.compile(sys.argv[3])
        for n, it in enumerate(stream):
            shown = it['text'] if it['kind'] == 'p' else f'<TABLE#{it["idx"]}>'
            if pat.search(shown):
                print(f"[{n:6d}] {fmt(it)}")
        return 0

    if mode == '--numheads':
        lo, hi = int(sys.argv[3]), int(sys.argv[4])
        for n in range(lo, min(hi, len(stream))):
            it = stream[n]
            if it['kind'] == 'p' and RE_ARAB.match(it['text']):
                print(f"[{n:6d}] {fmt(it)}")
        return 0

    if mode == '--table':
        want = int(sys.argv[3])
        for it in stream:
            if it['kind'] == 'table' and it['idx'] == want:
                print(f"TABLE#{want}  {it['rows']}行 x {it['cols']}列")
                for r, row in enumerate(it['grid'], 1):
                    cells = []
                    for c, cell in enumerate(row, 1):
                        if cell['m'] == ' ':
                            cells.append(f"{c}:{cell['t']}")
                        elif cell['m'] in ('h', 'v', 'x'):
                            cells.append(f"{c}:·{cell['m']}")
                    print(f"  R{r:<3d} " + ' | '.join(cells))
                return 0
        print(f'table {want} not found')
        return 1

    if mode == '--chapters':
        for n, it in enumerate(stream):
            if it['kind'] != 'p' or not it['text']:
                continue
            ol = it.get('outline')
            if ol is None or ol > 2:
                continue
            mark = '  <<<附注章' if is_note_chapter(it['text']) else ''
            print(f"[{n:6d}] H{ol + 1} {it.get('style_id', ''):<8} {it['text'][:70]}{mark}")
        return 0

    start = int(sys.argv[2])
    count = int(sys.argv[3]) if len(sys.argv) > 3 else 40
    for n in range(start, min(start + count, len(stream))):
        print(f"[{n:6d}] {fmt(stream[n])}")
    return 0


if __name__ == '__main__':
    sys.exit(main())
