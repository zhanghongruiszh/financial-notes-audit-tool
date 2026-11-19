# -*- coding: utf-8 -*-
"""
make_numbering_fixture.py — 造一份"编号坏掉"的测试文档，用来验证 numbering_fixer。

模拟三种真实故障：
  1. 双编号：去掉自动编号，但正文里手打「1、」「2、」，同时再挂一个自动编号
  2. 重启：让第 5 个之后的子项目挂到一个**新的**从 1 开始的列表上（显示 1、2、3、4）
  3. 无编号：其中一个子项目既没自动也没手工编号

用法：
    python tools/make_numbering_fixture.py <模板.docx> <输出.docx>
"""

import os
import sys

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import docx_numbering  # noqa: E402
import numbering_fixer  # noqa: E402


def body_paragraphs(doc):
    return [Paragraph(c, doc) for c in doc.element.body.iterchildren()
            if c.tag == qn('w:p')]


def text_len_without_numbering(p):
    """去掉自动编号后，正文文字的起始位置（这里直接用 run 拼）。"""
    return p.text


def main():
    src, dst = sys.argv[1], sys.argv[2]
    doc = Document(src)
    numbering_root = doc.part.numbering_part.element
    restart_num_id = numbering_fixer._build_single_level_list(numbering_root)

    paras = body_paragraphs(doc)
    resolver = docx_numbering.NumberingResolver(doc)

    # 找出所有 auto 编号为 ilvl=0 的段落（章标题），只在其后首个小节区间动手
    targets = []          # [(段落, 当前 auto 信息)]
    seen_chapter = False
    for p in paras:
        info = resolver.label(p)
        if not info:
            continue
        if info['ilvl'] == 0:
            # 章标题：只关心第一个"合并..."章
            if '主要项目附注' in p.text or '项目附注' in p.text:
                seen_chapter = True
            else:
                seen_chapter = False
            continue
        if info['ilvl'] == 1 and seen_chapter:
            targets.append(p)

    print(f'找到候选子项目段落 {len(targets)} 个')

    # 故障 1：前 2 个 → 去掉自动编号 + 手打编号 + 再挂自动编号（双编号）
    for i, p in enumerate(targets[:2], 1):
        prefix = f'{i}、'
        # 手打编号插到最前面
        if p.runs:
            p.runs[0].text = prefix + p.runs[0].text
        else:
            p.runs.append(p.add_run(prefix)) if hasattr(p, 'add_run') else None

    # 故障 2：第 5~8 个 → 挂到一个新的从 1 开始的列表（重启）
    for p in targets[4:8]:
        numbering_fixer._set_numpr(p, restart_num_id, ilvl=0)

    # 故障 3：第 10 个 → 去掉自动编号且不留手工（无编号）
    if len(targets) >= 10:
        pPr = targets[9]._p.find(qn('w:pPr'))
        if pPr is not None:
            numPr = pPr.find(qn('w:numPr'))
            if numPr is not None:
                pPr.remove(numPr)

    os.makedirs(os.path.dirname(os.path.abspath(dst)) or '.', exist_ok=True)
    doc.save(dst)
    print(f'已生成故障样本: {dst}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
