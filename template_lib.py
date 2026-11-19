# -*- coding: utf-8 -*-
"""
template_lib.py — 模板骨架库。

定位
----
模板**只作加速**，不是运行必需：没有模板照样能跑（自主识别），有模板时
识别更准、需要人工确认的地方更少。

骨架里存什么
------------
只存结构，不存原文（原始 .doc 共 18MB，不进包）：

    {
      "name": "一般企业模板6.4",
      "source": "一般企业模板6.4财务报表附注.doc",
      "chapters": [
        {"name": "...", "scope": "合并",
         "sub_items": [{"name": "货币资金", "seq": 1, "tables":
            [{"rows": 7, "cols": 3, "header": "项目|期末余额|期初余额",
              "roles": ["", "余额", "余额"]}]}]}
      ]
    }

存放位置
--------
程序同目录下的 `模板库/`，一个模板一个 JSON。用户可以直接把新的
.doc/.docx 丢进去，再在界面上点「重建骨架」——程序自己提取。

对外接口
--------
    library_dir()                      模板库目录（按需创建并首次播种）
    seed_if_empty()                    首次运行时把内置骨架放进去
    load_library()                     读取全部骨架
    build_skeleton(doc_path)           从 .doc/.docx 提取骨架
    save_skeleton(skeleton)            存成 JSON
    rebuild_from_docs(folder)          从原始文档批量重建
    best_match(model, library)         找最接近当前附注的模板
    roles_hint(skeleton, chapter, sub, table_seq, cols)
                                       给出某张表各列角色的建议
"""

import json
import os
import re
import shutil
import sys

LIBRARY_NAME = '模板库'
SKELETON_SUFFIX = '.skeleton.json'


# ---------------------------------------------------------------- 目录

def _app_dir():
    """程序所在目录（打包后是 exe 所在目录，源码运行时是项目根目录）。"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _bundled_dir():
    """打包进 exe 的内置骨架目录（未打包时就是项目里的 模板库）。"""
    meipass = getattr(sys, '_MEIPASS', None)
    if meipass:
        return os.path.join(meipass, LIBRARY_NAME)
    return os.path.join(_app_dir(), LIBRARY_NAME)


def library_dir():
    """用户可见的模板库目录：程序同目录下的 模板库/。"""
    path = os.path.join(_app_dir(), LIBRARY_NAME)
    os.makedirs(path, exist_ok=True)
    return path


def seed_if_empty():
    """首次运行把内置骨架拷到程序目录，之后以磁盘上的为准（用户可替换）。"""
    target = library_dir()
    existing = [f for f in os.listdir(target) if f.endswith(SKELETON_SUFFIX)]
    if existing:
        return target
    src = _bundled_dir()
    if os.path.isdir(src) and os.path.abspath(src) != os.path.abspath(target):
        for name in os.listdir(src):
            if name.endswith(SKELETON_SUFFIX):
                shutil.copy2(os.path.join(src, name), os.path.join(target, name))
    return target


# ---------------------------------------------------------------- 读写

def load_library(folder=None):
    folder = folder or seed_if_empty()
    out = []
    for name in sorted(os.listdir(folder)):
        if not name.endswith(SKELETON_SUFFIX):
            continue
        try:
            with open(os.path.join(folder, name), encoding='utf-8') as f:
                data = json.load(f)
            data['_file'] = name
            out.append(data)
        except Exception:
            continue
    return out


def save_skeleton(skeleton, folder=None):
    folder = folder or library_dir()
    name = skeleton.get('name') or '未命名模板'
    safe = re.sub(r'[\\/:*?"<>|]', '_', name)
    path = os.path.join(folder, safe + SKELETON_SUFFIX)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(skeleton, f, ensure_ascii=False, indent=1)
    return path


# ---------------------------------------------------------------- 提取

def build_skeleton(doc_path, work_dir=None):
    """从一份 .doc/.docx 提取骨架（子项目名单 + 每张表的行列与列角色）。"""
    import notes_parser
    import table_semantics as ts

    model = notes_parser.parse_notes(doc_path, work_dir=work_dir)
    by_sub = {}
    for tm in ts.analyze_document(model):
        by_sub.setdefault((tm.chapter, tm.sub_item), []).append(tm)

    chapters = []
    for ch in model.chapters:
        items = []
        for sub in ch.sub_items:
            tables = []
            for n, tm in enumerate(by_sub.get((ch.name, sub.name), []), 1):
                header = '|'.join(
                    (c.header_text or '').replace(' ', '')
                    for c in tm.columns if c.is_data)
                roles = [c.role or '' for c in tm.columns if c.is_data]
                tables.append({
                    'seq': n,
                    'rows': tm.ref.rows,
                    'cols': tm.ref.cols,
                    'header': header,
                    'roles': roles,
                })
            items.append({'name': sub.name, 'seq': sub.seq, 'tables': tables})
        chapters.append({'name': ch.name, 'scope': ch.scope,
                         'number': ch.number, 'sub_items': items})

    base = os.path.splitext(os.path.basename(doc_path))[0]
    return {
        'name': re.sub(r'(财务报表附注|附注|模板)$', '', base) or base,
        'source': os.path.basename(doc_path),
        'chapters': chapters,
    }


def rebuild_from_docs(folder=None, work_dir=None, log=print):
    """把模板库里（或其上级目录）的原始 .doc/.docx 全部重新提取成骨架。

    这是给"用户替换模板"用的：把新的 Word 模板放进模板库目录，
    点一下就能重建骨架，不用手写 JSON。
    """
    folder = folder or library_dir()
    docs = [os.path.join(folder, f) for f in sorted(os.listdir(folder))
            if f.lower().endswith(('.doc', '.docx')) and not f.startswith('~$')]
    if not docs:
        log(f'[模板] {folder} 里没有 .doc/.docx，无法重建')
        return []

    done = []
    for doc in docs:
        try:
            log(f'[模板] 提取骨架：{os.path.basename(doc)} ...')
            sk = build_skeleton(doc, work_dir=work_dir)
            path = save_skeleton(sk, folder)
            done.append(path)
            log(f'[模板]   子项目 {sum(len(c["sub_items"]) for c in sk["chapters"])} 个 '
                f'→ {os.path.basename(path)}')
        except Exception as exc:
            log(f'[模板]   失败：{exc}')
    return done


# ---------------------------------------------------------------- 匹配

def _sub_names(chapter):
    return {it.get('name') for it in chapter.get('sub_items', []) if it.get('name')}


def best_match(model, library):
    """找与当前附注最接近的模板：按子项目名重合度打分。

    不强制范围相同——单体附注常常和模板的「合并」章结构一致，
    所以同范围只给一点加分，而不是硬性条件。
    """
    if not library:
        return None, 0.0
    best, best_score = None, 0.0
    for sk in library:
        for ch in sk.get('chapters', []):
            names = _sub_names(ch)
            if not names:
                continue
            for mch in model.chapters:
                cur = {s.name for s in mch.sub_items}
                if not cur:
                    continue
                hit = len(cur & names)
                score = hit / max(len(cur | names), 1)
                if mch.scope == ch.get('scope'):
                    score += 0.05
                if score > best_score:
                    best, best_score = (sk, ch), score
    if best is None:
        return None, 0.0
    return best, min(best_score, 1.0)


def roles_hint(match, model, chapter_name, sub_name, table_seq, n_cols):
    """给出某张表各列角色的建议（模板里同位置的表）。

    match: best_match() 的第一个返回值 (skeleton, skeleton_chapter)
    返回长度 n_cols 的列表，元素为角色字符串或 None。
    """
    if not match:
        return [None] * n_cols
    sk, sk_chapter = match
    for item in sk_chapter.get('sub_items', []):
        if item.get('name') != sub_name:
            continue
        tables = item.get('tables') or []
        idx = table_seq - 1
        if idx < 0 or idx >= len(tables):
            return [None] * n_cols
        roles = tables[idx].get('roles') or []
        if len(roles) != n_cols:
            return [None] * n_cols
        return [r or None for r in roles]
    return [None] * n_cols
