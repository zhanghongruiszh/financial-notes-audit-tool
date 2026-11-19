# -*- coding: utf-8 -*-
"""
fingerprint.py — 项目结构指纹：把"这张表长什么样、怎么判的"存下来复用。

为什么要指纹库
--------------
空白模板只有表头，没有行；真实项目填过数，结构是"活的"。
同一个所的项目格式往往就那几种，跑几单就能把结构固定下来，
以后遇到同样结构直接套用，不用每次重新猜。

指纹存什么
----------
每条 = 一张表的结构签名 + 当时的判定结果：
    {
      "sub_item": "应收账款",
      "type": "账龄分析表",
      "rows": 7, "cols": 6,
      "headers": ["账龄", "期末余额 账面余额", ...],
      "row_labels": ["1年以内", "1至2年", ...],
      "roles": ["余额", "准备", ...],
      "period_kind": ["期末", "期初"],
      "total_rows": [7],
      "group_rows": [5]
    }

怎么用
------
    capture(model, tables)        从一个项目提取指纹
    save(project, prints, path)   写成 json（给测试人员导出用）
    load(path)                    读回
    match(tm, library)            按形状+表头找同结构，命中就复用判定
"""

import json
import os
import re

import table_semantics as ts

VERSION = 1


def _norm(text):
    """表头去掉空白和日期，只留结构词。"""
    s = str(text or '').strip()
    s = re.sub(r'\d{4}\s*[-年./]\s*\d{1,2}\s*[-月./]\s*\d{1,2}\s*日?', '<日期>', s)
    s = re.sub(r'\s+', '', s)
    return s


def capture_table(tm):
    """单张表的结构指纹。"""
    headers = [_norm(c.header_text) for c in tm.columns if c.is_data]
    roles = [c.role or '' for c in tm.columns if c.is_data]
    periods = []
    for c in tm.columns:
        if not c.is_data:
            continue
        periods.append(c.period_key or ('日期' if c.period else ''))
    labels = [(r.label or '').strip() for r in tm.rows
              if r.kind in ('detail', 'subtotal', 'group', 'total') and r.label]
    return {
        'sub_item': tm.sub_item,
        'chapter': tm.chapter,
        'type': getattr(tm, 'table_type', ''),
        'type_name': getattr(tm, 'table_type_name', ''),
        'rows': tm.ref.rows,
        'cols': tm.ref.cols,
        'headers': headers,
        'roles': roles,
        'period_kind': periods,
        'row_labels': labels[:40],
        'total_rows': [r.row for r in tm.rows if r.kind == 'total'],
        'sum_groups': _capture_sum_groups(tm),
        'subtotal_rows': [r.row for r in tm.rows if r.kind == 'subtotal'],
        'group_rows': [r.row for r in tm.rows if r.kind == 'group'],
    }


def tm_rows_label(tm, row_no):
    for r in tm.rows:
        if r.row == row_no:
            return r.label or ''
    return ''


def _capture_sum_groups(tm):
    """把这次"合计 = 哪几行"的判定也存下来，下次同结构直接照用。"""
    try:
        import rules
        checks = []
        rules._rule_v_sum(tm, checks)
        groups = {}
        grid = tm.ref.grid
        for chk in checks:
            if chk.rule_id != 'V-SUM':
                continue
            rows = [t.row for t in chk.terms]
            # 只存"验证通过"的成员集合；存疑的不要存，
            # 否则下次会把黄色存疑变成红色差异。
            try:
                import rules as _r
                target = _r.cell_number(grid, chk.target_row, chk.target_col)
                acc = sum(_r.row_sign(tm_rows_label(tm, r)) *
                          (_r.cell_number(grid, r, chk.target_col) or 0.0)
                          for r in rows)
                if target is not None and abs(acc - target) <= 0.01:
                    groups[str(chk.target_row)] = rows
            except Exception:
                continue
        return groups
    except Exception:
        return {}


def capture(model, tables, project=''):
    """整个项目的指纹。"""
    chapters = []
    if model is not None:
        chapters = [
            {'number': c.number, 'name': c.name, 'scope': c.scope,
             'sub_items': [s.name for s in c.sub_items]}
            for c in getattr(model, 'chapters', [])
        ]
    return {
        'version': VERSION,
        'project': project or os.path.basename(getattr(model, 'path', '') or ''),
        'chapters': chapters,
        'tables': [capture_table(t) for t in tables],
    }


def save(payload, path=None):
    if path is None:
        base = payload.get('project') or 'project'
        path = os.path.abspath(base + '.fingerprint.json')
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=1)
    return path


def load(path):
    with open(path, encoding='utf-8') as f:
        return json.load(f)


def library_dir():
    """指纹库存放目录：打包后放在 exe 旁边，源码运行时放在项目目录。"""
    import sys
    if getattr(sys, 'frozen', False):
        base = os.path.join(os.path.dirname(os.path.abspath(sys.executable)), '指纹库')
    else:
        base = os.path.join(os.path.dirname(os.path.abspath(__file__)), '指纹库')
    os.makedirs(base, exist_ok=True)
    return base


def seed_if_empty():
    """首次运行把内置指纹释放到程序目录（之后以磁盘上的为准）。"""
    import shutil
    import sys
    target = library_dir()
    if any(f.endswith('.fingerprint.json') for f in os.listdir(target)):
        return target
    mei = getattr(sys, '_MEIPASS', None)
    if not mei:
        return target
    src = os.path.join(mei, '指纹库')
    if os.path.isdir(src) and os.path.abspath(src) != os.path.abspath(target):
        for name in os.listdir(src):
            if name.endswith('.fingerprint.json'):
                shutil.copy2(os.path.join(src, name), os.path.join(target, name))
    return target


def load_library():
    base = library_dir()
    out = []
    for name in sorted(os.listdir(base)):
        if not name.endswith('.fingerprint.json'):
            continue
        try:
            out.append(load(os.path.join(base, name)))
        except Exception:
            continue
    return out


def install(payload):
    """把一份指纹装进指纹库。"""
    name = (payload.get('project') or 'project').replace('/', '_')
    path = os.path.join(library_dir(), name + '.fingerprint.json')
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=1)
    return path


def _signature(fp):
    return (fp.get('sub_item'), tuple(fp.get('headers') or ()),
            fp.get('rows'), fp.get('cols'))


def match(tm, library):
    """按「子项目 + 表头 + 行列数」找同结构的指纹。

    形状完全一致才认——和模板加速同样的原则，避免张冠李戴。
    """
    if not library:
        return None
    want = (_signature(capture_table(tm)))
    for payload in library:
        for fp in payload.get('tables', []):
            if _signature(fp) == want:
                return fp
    return None


def apply(fp, tm):
    """把指纹里的判定套到当前表上：补角色、补期间语义、标记合计行。"""
    if not fp:
        return 0
    tm.fingerprint = fp          # 供规则直接复用当时的合计成员
    data_cols = [c for c in tm.columns if c.is_data]
    filled = 0
    for col, role in zip(data_cols, fp.get('roles') or []):
        if role and not col.role:
            col.role = role
            if role in ts.RATIO_ROLES:
                col.kind = 'ratio'
            elif col.kind == 'text':
                col.kind = 'amount'
            filled += 1
    return filled
