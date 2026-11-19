# -*- coding: utf-8 -*-
"""
pipeline.py — 把整条流程串起来：导入 3 个文件 → 出 1 个成果文件夹。

输出结构（按需求确定）
--------------------
    <归档根目录>/
        <项目名>check_v1/
            1、审计报告附注-xxx.doc        ← 原始导入文件，备查
            2、审计报告报表-xxx.xlsx       ← 原始导入文件，备查
            TB-试算平衡表-xxx.xls          ← 原始导入文件，备查
            <项目名>check_v1.xlsx          ← 勾稽成果
        <项目名>check_v2/                 ← 二次核对自动递增版本号
            ...

TB 现阶段只作为备查文件归档；现金流量表与 TB 的逐格核对按需求留待后续。
"""

import os
import re
import shutil
import time

import cross_checks
import notes_parser
import report_reader as rr
import rules
import table_semantics as ts
import writer

# 文件名里用来切出项目名的分隔/噪声
RE_LEADING_NOISE = re.compile(r'^(TB|tb|报表|附注|\d|[·\-_\s、，,])+')
RE_PAREN = re.compile(r'[\(（].*?[\)）]')

# 文件名里这些片段不是项目名
DOC_TYPE_WORDS = ('审计报告', '附注', '报表', '财务报表', '试算平衡表', 'tb',
                  '主要项目附注', '资产负债表', '利润表', '现金流量表')
PERIOD_WORDS = re.compile(
    r'^(两年一期|三年一期|一年一期|年一期|申报|审定|合并|单体|\d{4}|[\d.]+\s*$)')


def _looks_like_project(seg):
    s = seg.strip(' -_·、')
    if len(s) < 2:
        return None
    low = s.lower()
    if any(w in low for w in DOC_TYPE_WORDS):
        return None
    if PERIOD_WORDS.match(s):
        return None
    return s


def infer_project_name(*paths):
    """从文件名推断项目名。

    常见命名：`2、审计报告报表-铁联化轻-两年一期.xlsx`
    → 按分隔符切段，丢掉"审计报告""报表""两年一期"这类非公司名片段，
    剩下的第一段就是项目名。推断不出返回空串，绝不瞎填。
    """
    for p in paths:
        if not p:
            continue
        base = os.path.splitext(os.path.basename(p))[0]
        clean = RE_PAREN.sub('', base)
        clean = RE_LEADING_NOISE.sub('', clean)
        for seg in re.split(r'[\-—_|]', clean):
            name = _looks_like_project(seg)
            if name:
                return name
    return ''


def _next_version(root, project_name):
    """扫描已有成果文件夹，返回下一个版本号。"""
    if not os.path.isdir(root):
        return 1
    pattern = re.compile(rf'^{re.escape(project_name)}check_v(\d+)$')
    used = [int(m.group(1)) for m in
            (pattern.match(d) for d in os.listdir(root)) if m]
    return max(used) + 1 if used else 1


def _preferred_input(src, suffix):
    """优先使用编号治理后的文件（如果存在）。

    用户可能先做了「Word附注编号重排」或「报表附注号修复」，那么勾稽就应该
    基于治理后的编号来做——否则刚修好的编号会被当成错误报出来。
    """
    if src and src.endswith(suffix):
        return src
    base = os.path.splitext(os.path.abspath(src))[0]
    cand = base + suffix
    return cand if os.path.exists(cand) else src


def run(notes_path, report_path, tb_path=None, archive_root='.',
        project_name=None, log=print):
    """执行全流程，返回结果字典。"""
    t0 = time.time()
    for label, p in (('报告附注', notes_path), ('财务报表', report_path)):
        if not p or not os.path.exists(p):
            raise FileNotFoundError(f'{label}文件不存在：{p}')
    if tb_path and not os.path.exists(tb_path):
        raise FileNotFoundError(f'TB 试算表文件不存在：{tb_path}')

    project = (project_name or '').strip() or \
        infer_project_name(report_path, notes_path, tb_path) or '未命名项目'
    version = _next_version(archive_root, project)
    out_dir = os.path.join(archive_root, f'{project}check_v{version}')
    os.makedirs(out_dir, exist_ok=True)
    log(f'[流程] 项目【{project}】第 {version} 版 → {out_dir}')

    # 0) 先把 .doc/.xls 统一转成 .docx/.xlsx
    #    （docx/xlsx 直接读取，不经过 Office；转换结果放在原文件旁边并复用）
    try:
        import convert_inputs
        converted = convert_inputs.convert_many(
            [p for p in (notes_path, report_path, tb_path) if p], log=log)
        mapping = {os.path.abspath(o): n for o, n, m in converted if n}
        notes_path = mapping.get(os.path.abspath(notes_path), notes_path)
        report_path = mapping.get(os.path.abspath(report_path), report_path)
        if tb_path:
            tb_path = mapping.get(os.path.abspath(tb_path), tb_path)
    except Exception as exc:
        log(f'[转换] 跳过格式转换：{exc}')

    # 1) 解析附注（若做过编号重排，用重排后的那份）
    notes_use = _preferred_input(notes_path, '_编号已规范.docx')
    report_use = _preferred_input(report_path, '_附注编号已修复.xlsx')
    used_note = ''
    if os.path.abspath(notes_use) != os.path.abspath(notes_path):
        used_note = f'附注采用编号规范后的：{os.path.basename(notes_use)}'
        log(f'[编号] {used_note}')
    if os.path.abspath(report_use) != os.path.abspath(report_path):
        msg = f'报表采用附注号修复后的：{os.path.basename(report_use)}'
        used_note = (used_note + '；' if used_note else '') + msg
        log(f'[编号] {msg}')

    log('[1/5] 解析报告附注 ...')
    model = notes_parser.parse_notes(notes_use, work_dir=os.path.join(out_dir, '_tmp'))
    log(f'      章 {len(model.chapters)} 个，子项目 {len(model.sub_items)} 个')

    # 2) 表格语义（有模板时用模板给识别兜底）
    log('[2/5] 识别表格结构 ...')
    hint = None
    try:
        import template_lib
        library = template_lib.load_library()
        if library:
            match, score = template_lib.best_match(model, library)
            if match:
                sk, sk_ch = match
                log(f'      模板库匹配：{sk.get("name")}（结构相似度 {score:.0%}）')
                hint = (lambda ch, sub, seq, n:
                        template_lib.roles_hint(match, model, ch, sub, seq, n))
            else:
                log(f'      模板库 {len(library)} 份，未匹配到相近结构，按自主识别')
    except Exception as exc:
        log(f'      模板库不可用（{exc}），按自主识别')

    tpl_stats = {}
    tables = ts.analyze_document(model, hint=hint, stats=tpl_stats)

    # 指纹库：同结构（子项目+表头+行列数完全一致）直接套用上次的判定
    try:
        import fingerprint
        lib = fingerprint.load_library()
        if lib:
            matched = filled = 0
            for tb in tables:
                fp = fingerprint.match(tb, lib)
                if fp:
                    matched += 1
                    filled += fingerprint.apply(fp, tb)
            if matched:
                log(f'      指纹库命中 {matched} 张表（库内 {len(lib)} 个项目），'
                    f'补全列角色 {filled} 处')
    except Exception as exc:
        log(f'      指纹库不可用（{exc}）')
    log(f'      表 {len(tables)} 张'
        + (f'，模板补全列角色 {tpl_stats.get("template_roles", 0)} 处'
           if tpl_stats.get('template_roles') else ''))

    # 3) 附注内部勾稽
    log('[3/5] 生成并计算附注内部勾稽 ...')
    checks = rules.build_checks(tables)
    by_uid = rules.map_tables(tables)
    results = rules.evaluate_all(checks, by_uid)
    summary = rules.summarize(results)
    log(f'      勾稽 {len(checks)} 条：{summary}')

    # 4) 报表勾稽
    log('[4/5] 读取三张报表并做跨表勾稽 ...')
    report = rr.read_report(report_use)
    log(f'      识别到报表 {len(report.sheets)} 张：'
        + '、'.join(s.name for s in report.sheets))
    import openpyxl
    wb_src = openpyxl.load_workbook(report_use, data_only=True)

    def _get_value(sheet_name, row, col):
        if sheet_name not in wb_src.sheetnames:
            return None
        v = wb_src[sheet_name].cell(row, col).value
        return float(v) if isinstance(v, (int, float)) else None

    cross = cross_checks.build_cross_checks(
        model, tables, report, get_value=_get_value)
    report_values = {}
    for cc in cross:
        if cc.sheet and cc.rows and cc.cols and cc.sheet in wb_src.sheetnames:
            report_values[(cc.sheet, cc.rows, cc.cols)] = \
                wb_src[cc.sheet].cell(cc.rows, cc.cols).value
    for cc in cross:
        cross_checks.evaluate(cc, report_values)
    log(f'      跨表勾稽 {len(cross)} 条')

    # 4a) 报表附注号 与 附注实际编号 的一致性
    log('[4a/5] 核对报表「附注」列与附注实际编号 ...')
    import note_number_checks as nn_mod
    nn_results = nn_mod.build_note_no_checks(report_use, model)
    if nn_results:
        log(f'      附注号勾稽 {len(nn_results)} 条：{nn_mod.summarize(nn_results)}')

    # 4c) 报表项目在附注里有没有披露（只汇总，不下结论）
    log('[4c/5] 核对报表项目是否在附注中披露 ...')
    import disclosure_checks as dc_mod
    disclosure = dc_mod.build_disclosure_checks(report_use, report, model)
    log(f'      未披露候选 {len(disclosure)} 条（供人工判断）')

    # 4d) 附注内部跨表 / 跨子项目（如 现金余额 vs 货币资金合计）
    log('[4d/5] 核对附注内部跨表与上下文相关项目 ...')
    import note_cross_checks as nx_mod
    note_cross = nx_mod.build(model, tables)
    if note_cross:
        log(f'      附注内跨表 {len(note_cross)} 条：{nx_mod.summarize(note_cross)}')

    # 5) TB 现金流量表核对
    tb_results = []
    tb_xlsx = None
    if not tb_path:
        log('[4b/5] 未导入 TB，跳过现金流量表子项核对（成果表会标注）')
    if tb_path:
        log('[4b/5] 读取 TB 现金流量表并逐项核对 ...')
        import tb_reader
        import tb_checks as tb_mod
        tmp_work = os.path.join(out_dir, '_tmp')
        tb_sheets = tb_reader.read_tb(tb_path, work_dir=tmp_work)
        if tb_sheets:
            log(f'      识别到 TB 现金流量表：{"、".join(tb_sheets)}')
            tb_xlsx = tb_reader.ensure_xlsx(tb_path, work_dir=tmp_work)
            tb_results = tb_mod.build_tb_checks(model, tables, tb_sheets)
            tb_mod.evaluate(tb_results, tb_xlsx)
            from collections import Counter as _C
            log(f'      TB 核对 {len(tb_results)} 条：{dict(_C(c.status for c in tb_results))}')
        else:
            log('      没识别到 TB 现金流量表（工作表名需含 CF）')

    # 6) 生成成果表
    log('[5/5] 生成成果表 ...')
    out_xlsx = os.path.join(out_dir, f'{project}check_v{version}.xlsx')
    writer.write_workbook(out_xlsx, model, tables, checks, results,
                          cross=cross, report=report,
                          tb_checks=tb_results or None, tb_xlsx=tb_xlsx,
                          note_no_checks=nn_results or None,
                          disclosure=disclosure,
                          note_cross=note_cross or None,
                          source_note=used_note, log=log)

    # 归档导入的三个原始文件，以及（若存在）编号治理产出的两个文件
    archived = []
    seen = set()

    def _archive(src):
        if not src or not os.path.exists(src):
            return
        name = os.path.basename(src)
        if name in seen:
            return
        dst = os.path.join(out_dir, name)
        if os.path.abspath(src) != os.path.abspath(dst):
            shutil.copy2(src, dst)
        archived.append(name)
        seen.add(name)

    for src in (notes_path, report_path, tb_path):
        _archive(src)

    # 编号治理产物：Word 编号规范后的附注、报表附注号修复后的报表
    for src in (notes_path, report_path):
        if not src:
            continue
        base = os.path.splitext(os.path.abspath(src))[0]
        for suffix in ('_编号已规范.docx', '_附注编号已修复.xlsx'):
            _archive(base + suffix)

    # 清掉临时转换目录
    tmp_dir = os.path.join(out_dir, '_tmp')
    if os.path.isdir(tmp_dir):
        shutil.rmtree(tmp_dir, ignore_errors=True)

    elapsed = time.time() - t0
    log(f'[流程] 完成，用时 {elapsed:.1f}s')
    log(f'[流程] 成果文件：{out_xlsx}')
    log(f'[流程] 归档：{"、".join(archived)}')

    return {
        'project': project,
        'version': version,
        'out_dir': out_dir,
        'out_xlsx': out_xlsx,
        'archived': archived,
        'stats': dict(summary),
        'checks': len(checks),
        'cross': len(cross),
        'elapsed': elapsed,
    }


def _main():
    import sys
    if len(sys.argv) < 3:
        print('用法: python pipeline.py <附注.doc> <报表.xlsx> [TB.xls] [归档目录]')
        return 2
    notes, report = sys.argv[1], sys.argv[2]
    tb = sys.argv[3] if len(sys.argv) > 3 else None
    root = sys.argv[4] if len(sys.argv) > 4 else '_work/out'
    res = run(notes, report, tb, archive_root=root)
    print(res)
    return 0


if __name__ == '__main__':
    raise SystemExit(_main())
