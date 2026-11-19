# -*- coding: utf-8 -*-
"""
make_renumbered_fixture.py — 造一份"编号已被重排"的附注，用于验证勾稽口径。

把附注里指定序号的子项目标题改成别的编号（模拟重排后编号变化），
存成 `原文件名_编号已规范.docx`，pipeline 会优先采用它。

用法：
    python tools/make_renumbered_fixture.py <源docx> <输出docx> <旧编号> <新编号> [...]
"""

import os
import re
import sys

from docx import Document

RE_HEAD = re.compile(r'^(\d+)(\s*[、.．]\s*)(.*)$')


def main():
    src, dst = sys.argv[1], sys.argv[2]
    pairs = []
    rest = sys.argv[3:]
    for i in range(0, len(rest) - 1, 2):
        pairs.append((rest[i], rest[i + 1]))
    if not pairs:
        print('至少给一组 旧编号 新编号')
        return 2

    doc = Document(src)
    changed = 0
    for p in doc.paragraphs:
        text = p.text.strip()
        m = RE_HEAD.match(text)
        if not m:
            continue
        old = m.group(1)
        hit = next((new for o, new in pairs if o == old), None)
        if hit is None:
            continue
        # 只在标题段改（标题段一般没有句号、长度较短）
        if len(text) > 30:
            continue
        new_text = f'{hit}{m.group(2)}{m.group(3)}'
        if p.runs:
            # 把整段文字塞进第一个 run，其余清空，避免编号被拆在多个 run 里
            p.runs[0].text = new_text
            for r in p.runs[1:]:
                r.text = ''
        changed += 1
        print(f'  「{text}」→「{new_text}」')

    os.makedirs(os.path.dirname(os.path.abspath(dst)) or '.', exist_ok=True)
    doc.save(dst)
    print(f'共改 {changed} 处，已保存 {dst}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
