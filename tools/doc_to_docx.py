# -*- coding: utf-8 -*-
"""
doc_to_docx.py — 用 Word 一次性把 .doc 另存为 .docx，后续全部离线解析。

逐单元格走 COM 极慢（2MB 文档 252 张表要 29 分钟）；转成 .docx 后
用 python-docx 读，同一份文档不到 1 秒，且能拿到完整的网格坐标与合并信息。

用法：
    python tools/doc_to_docx.py <输出目录> <文件1> [文件2 ...]
    python tools/doc_to_docx.py _docx "附注模板"
"""

import os
import sys
import time
import traceback

import win32com.client as win32

WD_FORMAT_XML_DOCUMENT = 16


def convert_batch(out_dir, paths, log=print):
    os.makedirs(out_dir, exist_ok=True)
    word = win32.DispatchEx('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        word.AutomationSecurity = 3
    except Exception:
        pass

    results = []
    try:
        for i, src in enumerate(paths, 1):
            base = os.path.splitext(os.path.basename(src))[0]
            dst = os.path.join(out_dir, base + '.docx')
            if os.path.exists(dst) and os.path.getmtime(dst) >= os.path.getmtime(src):
                log(f'[{i}/{len(paths)}] skip (up to date): {base}')
                results.append((src, dst, True))
                continue
            t0 = time.time()
            doc = None
            try:
                doc = word.Documents.Open(os.path.abspath(src), ConfirmConversions=False,
                                          ReadOnly=True, AddToRecentFiles=False)
                doc.SaveAs2(os.path.abspath(dst), FileFormat=WD_FORMAT_XML_DOCUMENT)
                doc.Close(False)
                doc = None
                results.append((src, dst, True))
                log(f'[{i}/{len(paths)}] OK {time.time() - t0:.1f}s -> {dst}')
            except Exception:
                log(f'[{i}/{len(paths)}] FAIL {base}\n{traceback.format_exc()}')
                results.append((src, dst, False))
                if doc is not None:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
    finally:
        try:
            word.Quit()
        except Exception:
            pass
    return results


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        return 2
    out_dir = sys.argv[1]
    targets = sys.argv[2:]
    paths = []
    for t in targets:
        if os.path.isdir(t):
            for f in sorted(os.listdir(t)):
                if f.lower().endswith(('.doc', '.docx')) and not f.startswith('~$'):
                    paths.append(os.path.join(t, f))
        else:
            paths.append(t)

    logf = open(os.path.join(out_dir if os.path.isdir(out_dir) else '.', '_convert.log'),
                'w', encoding='utf-8')

    def log(msg):
        print(msg)
        logf.write(str(msg) + '\n')
        logf.flush()

    results = convert_batch(out_dir, paths, log)
    ok = sum(1 for _, _, s in results if s)
    log(f'DONE ok={ok}/{len(results)}')
    logf.close()
    return 0 if ok == len(results) else 1


if __name__ == '__main__':
    sys.exit(main())
