# -*- coding: utf-8 -*-
"""
Word 审计报告附注自动提取工具 (Word-to-Excel Extractor)

从审计报告 Word 文档（.doc / .docx）中定位「主要项目附注」章节，把该章内
**所有段落与表格** 按文档物理顺序铺进 Excel，得到可供勾稽的基础表 check.xlsx。

与旧版的区别
------------
旧版只写段落、**完全不写表格**，导致金额全部丢失（一份真实附注只剩 94 行标题、
0 行数据），下游勾稽引擎随即崩溃。本版：

  - 段落：一个段落 = 一行，文字放 A 列（空段落也占一行，保持行号与文档一致）
  - 表格：一个表格行 = 一行，一个单元格 = 一列；合并区只写左上角，其余留空，
          所以列位不会塌陷，能和原表逐格对上
  - 数值：`7,921,661.42` -> 7921661.42；`100.00%` -> 1.0；
          `2026年9月30日` -> 日期；`-` 保留为文字（表示"无"，不是 0）
  - 章节：靠「关键词 + 大纲级别」定位，不写死章节号，也不写死科目

不再依赖逐单元格的 Word COM 调用：先用 Word 把 .doc 转成 .docx，再离线解析，
2MB 的模板从 29 分钟降到秒级。
"""

import os
import sys

import openpyxl

import docx_ir

# A 列宽度，便于人工在 Excel 里看
COL_A_WIDTH = 42


def extract_notes_from_word(doc_filename=None, out_excel='extracted_notes.xlsx', verbose=True):
    """把 Word 附注的主要项目章节提取到 Excel。返回输出文件路径。"""
    def log(msg):
        if verbose:
            print(msg)

    if not doc_filename:
        docs = [f for f in os.listdir('.')
                if f.lower().endswith(('.doc', '.docx')) and not f.startswith('~$')]
        if not docs:
            raise FileNotFoundError('未在当前目录下找到 Word 报告文件 (.doc / .docx)')
        doc_filename = docs[0]

    log(f'[提取] 源文件: {doc_filename}')
    docx_path = docx_ir.ensure_docx(doc_filename)
    log(f'[提取] 归一为 docx: {os.path.basename(docx_path)}')

    stream = docx_ir.load_stream(docx_path)
    chapters = docx_ir.find_note_chapters(stream)
    if not chapters:
        raise ValueError(
            '未在文档中找到「主要项目附注」章节。\n'
            '已识别的章标题关键词：' + '、'.join(docx_ir.NOTE_CHAPTER_KEYS) + '\n'
            '请确认该章节标题文字，或确认文档是否带大纲级别（Word 中的「标题 1/标题 2」样式）。'
        )

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Sheet1'
    ws.column_dimensions['A'].width = COL_A_WIDTH

    row_idx = 1
    summary = []
    for ch in chapters:
        items = docx_ir.find_sub_items(stream, ch)
        log(f"[提取] 章<{ch['name']}>  范围[{ch['start']},{ch['end']})  "
            f"子项目 {len(items)} 个")
        rows = docx_ir.iter_chapter_rows(stream, ch)
        for r in rows:
            for c, val in enumerate(r['cells'], 1):
                if val is not None:
                    ws.cell(row_idx, c, val)
            row_idx += 1
        summary.append((ch, items, len(rows)))

    os.makedirs(os.path.dirname(os.path.abspath(out_excel)) or '.', exist_ok=True)
    wb.save(out_excel)

    total_tables = sum(1 for x in stream if x['kind'] == 'table')
    log(f'[提取] 完成：{len(chapters)} 章、{total_tables} 张表、共 {row_idx - 1} 行')
    log(f'[提取] 已保存: {out_excel}')
    return out_excel


def _main():
    target = sys.argv[1] if len(sys.argv) > 1 else None
    out = sys.argv[2] if len(sys.argv) > 2 else None

    if target and os.path.isfile(target):
        doc_path = target
        out_path = out or os.path.join(os.path.dirname(os.path.abspath(doc_path)), 'check.xlsx')
    else:
        folder = target or '.'
        docs = [f for f in os.listdir(folder)
                if f.lower().endswith(('.doc', '.docx')) and not f.startswith('~$')]
        if not docs:
            print(f'在目录 {folder} 中未找到 Word 文档')
            return 1
        doc_path = os.path.join(folder, docs[0])
        out_path = out or os.path.join(folder, 'check.xlsx')

    extract_notes_from_word(doc_path, out_path)
    return 0


if __name__ == '__main__':
    sys.exit(_main())
