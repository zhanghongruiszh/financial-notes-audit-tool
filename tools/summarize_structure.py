# -*- coding: utf-8 -*-
"""
summarize_structure.py — 汇总 analyze_docx 的 JSON，给出各模板的章节/子项目/表格画像。

用法：
    python tools/summarize_structure.py _docx
"""

import json
import os
import re
import sys
from collections import Counter


CN_NUM = '一二三四五六七八九十百'
RE_CN_HEAD = re.compile(rf'^[{CN_NUM}]+[、．.]')
RE_ARAB_HEAD = re.compile(r'^(\d+)[、．.]\s*(\S.*)$')
RE_PAREN = re.compile(r'^[（(]\s*\d+\s*[)）]')

NOTE_CHAPTER_KEYS = ('主要项目附注', '主要项目注释', '财务报表项目注释',
                     '财务报表主要项目', '报表项目注释', '主要报表项目')


def is_note_chapter(text):
    return any(k in text for k in NOTE_CHAPTER_KEYS)


def load(path):
    with open(path, encoding='utf-8') as f:
        return json.load(f)


def summarize(path):
    res = load(path)
    stream = res['stream']
    out = []
    out.append('=' * 100)
    out.append(f"FILE: {os.path.basename(path)}   paragraphs={res['paragraph_count']} tables={res['table_count']}")

    # 大纲级别分布
    lv = Counter()
    styles = Counter()
    for it in stream:
        if it['kind'] == 'p' and it['text']:
            lv[it.get('outline')] += 1
            styles[(it.get('style_id'), it.get('outline'))] += 1
    out.append(f"大纲级别分布 (None=无): {dict(sorted(lv.items(), key=lambda x: (x[0] is None, x[0])))}")
    out.append('样式分布 top10: ' + str(styles.most_common(10)))

    # 章（outline 0 或 无级别但以"一、"开头且很短）
    chapters = []
    for idx, it in enumerate(stream):
        if it['kind'] != 'p' or not it['text']:
            continue
        t = it['text']
        if it.get('outline') == 0 or (RE_CN_HEAD.match(t) and len(t) <= 40 and it.get('outline') in (0, 1, None)):
            chapters.append((idx, t, it.get('outline'), it.get('style_id')))
    out.append(f"--- 候选章标题 {len(chapters)} 个 ---")
    for idx, t, ol, sid in chapters:
        flag = ' <<< 附注章' if is_note_chapter(t) else ''
        out.append(f"  [{idx:6d}] H{'' if ol is None else ol + 1} {sid:<10} {t[:60]}{flag}")

    # 在有"附注章"的区间内找子项目标题
    note_spans = []
    for i, (idx, t, ol, sid) in enumerate(chapters):
        if not is_note_chapter(t):
            continue
        end = chapters[i + 1][0] if i + 1 < len(chapters) else len(stream)
        note_spans.append((idx, end, t))
    out.append(f"--- 附注章区间 {len(note_spans)} 个 ---")
    for start, end, t in note_spans:
        subs = []
        for it in stream[start + 1:end]:
            if it['kind'] != 'p' or not it['text']:
                continue
            m = RE_ARAB_HEAD.match(it['text'])
            if not m:
                continue
            ol = it.get('outline')
            if ol is None or ol >= 6:
                continue
            subs.append((it['text'][:40], ol, it.get('style_id')))
        tcount = sum(1 for it in stream[start + 1:end] if it['kind'] == 'table')
        out.append(f"  章<{t[:40]}>  子项目={len(subs)}  表={tcount}")
        lvl_counter = Counter(ol for _, ol, _ in subs)
        out.append(f"    子项目大纲级别分布: {dict(sorted(lvl_counter.items()))}")
        for s in subs[:70]:
            out.append(f"      H{s[1] + 1} {s[2]:<8} {s[0]}")
        if len(subs) > 70:
            out.append(f"      ... 另 {len(subs) - 70} 个")

    # 表格形态画像
    shapes = Counter()
    big = []
    for it in stream:
        if it['kind'] != 'table':
            continue
        shapes[(it['rows'], it['cols'])] += 1
        if it['rows'] * it['cols'] >= 60:
            big.append((it['idx'], it['rows'], it['cols']))
    out.append(f"--- 表格形态 top15 (行,列):次数 ---")
    for shape, n in shapes.most_common(15):
        out.append(f"  {shape[0]}x{shape[1]}: {n}")
    out.append(f"  最大表格: {sorted(big, key=lambda x: -x[1] * x[2])[:8]}")
    return '\n'.join(out)


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else '_work/docx'
    files = sorted(f for f in os.listdir(folder) if f.startswith('_a_') and f.endswith('.json'))
    chunks = []
    for f in files:
        chunks.append(summarize(os.path.join(folder, f)))
    text = '\n\n'.join(chunks)
    out_path = os.path.join(folder, '_summary.txt')
    open(out_path, 'w', encoding='utf-8').write(text)
    print(f'written {out_path} ({len(text)} chars)')
    return 0


if __name__ == '__main__':
    sys.exit(main())
