# -*- coding: utf-8 -*-
"""
diagnostics.py — 导出"测试包"，把出问题时需要的信息一次打包。

测试人员遇到报错、或者大量表没认出来时，点一下就能生成一个压缩包发回来。
包里有：
    说明.txt            这个包是什么、怎么来的
    运行环境.txt         程序版本、系统、时间
    日志.txt            界面上的运行日志
    结构报告.csv         每张表识别成什么样（类型/行列/列角色/合计行/是否存疑）
    结构报告.txt         同上，按表排版，便于直接看
    <项目>.fingerprint.json   结构指纹（不含金额）
    输入文件/            三个原始文件的副本（含业务数据，是否发送由你决定）
"""

import csv
import datetime
import os
import platform
import shutil
import sys
import zipfile

APP_NAME = '财务报表附注自动化勾稽系统'
APP_VERSION = '测试版'


def _env_text(paths):
    lines = [f'{APP_NAME}（{APP_VERSION}）',
             f'导出时间：{datetime.datetime.now():%Y-%m-%d %H:%M:%S}',
             f'操作系统：{platform.platform()}',
             f'Python：{sys.version.split()[0]}',
             f'打包运行：{getattr(sys, "frozen", False)}', '']
    for label, p in paths:
        lines.append(f'{label}：{p}')
        if p and os.path.exists(p):
            lines.append(f'    大小：{os.path.getsize(p)} 字节')
        else:
            lines.append('    （未提供或不存在）')
    return '\n'.join(lines)


def _table_rows(tables):
    out = []
    for t in tables:
        rows = [r for r in t.rows if r.kind in ('detail', 'subtotal', 'total', 'group')]
        totals = [f'R{r.row} {r.label}' for r in t.rows if r.kind == 'total']
        groups = [f'R{r.row} {r.label}' for r in t.rows if r.kind == 'group']
        cols = [f'{c.index + 1}:{c.header_text[:18]}({c.role or "-"}/{c.kind})'
                for c in t.columns if c.is_data]
        unclear = '结构存疑' if any('结构存疑' in w for w in t.warnings) else ''
        out.append({
            '表号': t.ref.idx,
            '子项目': t.sub_item,
            '类型': getattr(t, 'table_type_name', ''),
            '行数': t.ref.rows,
            '列数': t.ref.cols,
            '列角色': ' | '.join(cols),
            '合计行': ' | '.join(totals),
            '其中行': ' | '.join(groups),
            '明细行数': len(rows),
            '警告': '；'.join(t.warnings) or unclear,
        })
    return out


def build(notes, report, tb, log_text, out_path, project=''):
    """生成测试包（zip），返回压缩包路径。"""
    import notes_parser
    import table_semantics as ts

    stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    base = f'{project or "项目"}_测试包_{stamp}'
    work = os.path.join(os.path.dirname(os.path.abspath(out_path)), base)
    os.makedirs(work, exist_ok=True)

    tables = []
    err = ''
    try:
        model = notes_parser.parse_notes(notes, work_dir=os.path.join(work, '_parse'))
        tables = ts.analyze_document(model)
    except Exception as exc:          # noqa: BLE001
        err = f'解析附注失败：{exc}'

    with open(os.path.join(work, '运行环境.txt'), 'w', encoding='utf-8') as f:
        f.write(_env_text([('报告附注', notes), ('财务报表', report), ('TB 试算表', tb)]))

    with open(os.path.join(work, '日志.txt'), 'w', encoding='utf-8') as f:
        f.write(log_text or '（无日志）')
        if err:
            f.write('\n\n' + err)

    rows = _table_rows(tables)
    if rows:
        with open(os.path.join(work, '结构报告.csv'), 'w', encoding='utf-8-sig',
                  newline='') as f:
            w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
            w.writeheader()
            w.writerows(rows)

        lines = []
        for r in rows:
            lines.append(f"[表#{r['表号']}] {r['子项目']}　{r['类型']}　"
                         f"{r['行数']}行×{r['列数']}列　明细{r['明细行数']}行")
            if r['列角色']:
                lines.append('    列：' + r['列角色'])
            if r['合计行']:
                lines.append('    合计：' + r['合计行'])
            if r['警告']:
                lines.append('    ⚠ ' + r['警告'])
            lines.append('')
        with open(os.path.join(work, '结构报告.txt'), 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines))

        try:
            import fingerprint
            payload = fingerprint.capture(None, tables, project=project)
            fingerprint.save(payload, os.path.join(work, f'{project or "项目"}.fingerprint.json'))
        except Exception:
            pass

    inputs = os.path.join(work, '输入文件')
    os.makedirs(inputs, exist_ok=True)
    for p in (notes, report, tb):
        if p and os.path.exists(p):
            try:
                shutil.copy2(p, os.path.join(inputs, os.path.basename(p)))
            except Exception:
                pass

    with open(os.path.join(work, '说明.txt'), 'w', encoding='utf-8') as f:
        f.write(
            '这是一份勾稽程序的"测试包"，用于排查问题。\n\n'
            '包含内容：\n'
            '  · 运行环境.txt   程序版本与文件路径\n'
            '  · 日志.txt       界面上的运行日志（含报错）\n'
            '  · 结构报告.csv / .txt  每张表被识别成什么样\n'
            '  · *.fingerprint.json   表结构指纹（不含金额）\n'
            '  · 输入文件/      三个原始文件的副本\n\n'
            '注意：输入文件里含业务数据，是否发送请自行判断；\n'
            '      如不方便，可只发结构报告和日志。\n')

    log_copy = os.path.join(work, '_parse')
    if os.path.isdir(log_copy):
        shutil.rmtree(log_copy, ignore_errors=True)

    zip_path = work + '.zip'
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, _dirs, files in os.walk(work):
            for name in files:
                full = os.path.join(root, name)
                z.write(full, os.path.relpath(full, work))
    shutil.rmtree(work, ignore_errors=True)
    return zip_path
