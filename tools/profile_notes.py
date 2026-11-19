# -*- coding: utf-8 -*-
"""
profile_notes.py — 验证「章 / 子项目」识别规则在全部文档上的通用性。

规则（不依赖样式ID、不依赖编号）：
  1) 章 = 文本命中附注关键词、且带大纲级别的段落；取其大纲级别 L。
  2) 章区间 = 从该段落到下一个大纲级别 <= L 的标题为止。
  3) 子项目级别 = 该区间内出现的、最小的 (级别 > L) 的大纲级别。
  4) 子项目 = 区间内处于该级别的段落。

用法：
    python tools/profile_notes.py _docx
"""

import json
import os
import re
import sys

NOTE_KEYS = ('主要项目附注', '主要项目注释', '财务报表项目注释', '财务报表项目附注',
             '报表项目附注', '报表项目注释', '报表主要项目附注')
RE_ARAB = re.compile(r'^\s*(\d+)\s*[、．.]\s*(\S.*)$')
RE_PAREN = re.compile(r'^[（(]\s*\d+\s*[)）]')


def is_note_chapter(text):
    if not any(k in text for k in NOTE_KEYS):
        return False
    # 排除"合并范围的变更""重要会计政策"等无关命中
    bad = ('会计政策', '变更', '差错', '编制基础')
    return not any(b in text for b in bad)


def analyze(path):
    res = json.load(open(path, encoding='utf-8'))
    stream = res['stream']

    # 带大纲级别的段落索引
    heads = [(n, it) for n, it in enumerate(stream)
             if it['kind'] == 'p' and it['text'] and it.get('outline') is not None]

    out = [f"FILE: {os.path.basename(path)}  paragraphs={res['paragraph_count']} "
           f"tables={res['table_count']}"]

    for pos, (n, it) in enumerate(heads):
        if not is_note_chapter(it['text']):
            continue
        L = it['outline']
        # 区间终点：下一个级别 <= L 的标题
        end = len(stream)
        for n2, it2 in heads[pos + 1:]:
            if it2['outline'] <= L:
                end = n2
                break

        # 子项目级别 = 区间内最小的 > L 的级别
        levels = [it2['outline'] for n2, it2 in heads if n < n2 < end and it2['outline'] > L]
        sub_level = min(levels) if levels else None

        subs = [(n2, it2) for n2, it2 in heads
                if n < n2 < end and sub_level is not None and it2['outline'] == sub_level]

        tcount = sum(1 for x in stream[n:end] if x['kind'] == 'table')
        numbered = sum(1 for _, s in subs if RE_ARAB.match(s['text']))

        out.append(f"  章: <{it['text']}>  outline={L}  区间=[{n},{end})  表={tcount}")
        out.append(f"    子项目级别={sub_level}  数量={len(subs)}  带数字编号={numbered}")
        for n2, s in subs[:25]:
            m = RE_ARAB.match(s['text'])
            no = f"#{m.group(1)}" if m else '  -'
            out.append(f"      [{n2:6d}] {no:>4} {s['text'][:50]}")
        if len(subs) > 25:
            out.append(f"      ... 另 {len(subs) - 25} 个")
        out.append('')
    return '\n'.join(out)


def main():
    folder = sys.argv[1] if len(sys.argv) > 1 else '_work/docx'
    files = sorted(f for f in os.listdir(folder) if f.startswith('_a_') and f.endswith('.json'))
    text = '\n'.join(analyze(os.path.join(folder, f)) for f in files)
    out_path = os.path.join(folder, '_notes_profile.txt')
    open(out_path, 'w', encoding='utf-8').write(text)
    print(f'written {out_path}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
