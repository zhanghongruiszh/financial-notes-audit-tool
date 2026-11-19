# -*- coding: utf-8 -*-
"""
learn_rules.py — 把"猜中的结果"沉淀成模糊关键字规则。

设计原则（对应 docs/例外规则登记表.md）：
  1. 先按**已学到的关键字**匹配（快、可解释）；
  2. 匹配不上再按**数值**猜（慢，但一定能对上一个）；
  3. 猜中之后，从**表头/行名的语义**里抽出关键字，记进规则库；
  4. 下次同一个科目优先用关键字，不再靠猜。

规则库落在 `规则库/learned_rules.json`，就是一个可读的字典，
人工可以直接改、也可以删掉重新学：

    {
      "row:租赁负债":      {"keyword": "净额",  "hits": 3},
      "row:长期股权投资":  {"keyword": "合计",  "hits": 2},
      "col:应收账款":      {"keyword": "账面余额", "hits": 1}
    }

关键字都取自 KEYWORD_POOL，是"模糊片段"而不是整行文字，
所以「租赁负债净额」「租赁负债账面净额」都能命中「净额」。
"""

import json
import os
import re

# 语义关键字池：按"越具体越优先"的顺序排列
KEYWORD_POOL = [
    '账面价值', '账面净值', '账面余额', '账面原值',
    '净额', '净值', '原值', '原价',
    '累计折旧', '累计摊销', '减值准备', '减值',
    '期初余额', '期末余额', '年初余额', '年末余额',
    '本期发生额', '上期发生额', '本年增加', '本年减少',
    '合计', '小计', '总计', '余额', '金额',
]

_CACHE = None


def library_dir():
    """规则库目录：打包后放在 exe 同级，源码运行放项目根目录。"""
    import sys
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(os.path.abspath(sys.executable))
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    d = os.path.join(base, '规则库')
    os.makedirs(d, exist_ok=True)
    return d


def _path():
    return os.path.join(library_dir(), 'learned_rules.json')


def load():
    global _CACHE
    if _CACHE is not None:
        return _CACHE
    try:
        with open(_path(), encoding='utf-8') as f:
            _CACHE = json.load(f)
    except Exception:
        _CACHE = {}
    return _CACHE


def save():
    try:
        with open(_path(), 'w', encoding='utf-8') as f:
            json.dump(_CACHE or {}, f, ensure_ascii=False, indent=1)
        return True
    except Exception:
        return False


# 只学"有区分度"的关键字：小计/合计/金额/余额 这类太泛，
# 学了反而会到处误命中
LEARNABLE = ('账面价值', '账面净值', '账面余额', '账面原值',
             '净额', '净值', '原值', '原价',
             '累计折旧', '累计摊销', '减值准备')


def extract_keyword(text):
    """从一段表头/行名里抽出语义关键字（模糊片段，不是全名）。"""
    if not text:
        return None
    t = re.sub(r'\s+', '', str(text))
    for kw in LEARNABLE:
        if kw in t:
            return kw
    return None


def learn(kind, name, text):
    """记下：某个科目，代表行/列的文字里出现了哪个关键字。"""
    if not name or not text:
        return None
    kw = extract_keyword(text)
    if not kw:
        return None
    lib = load()
    key = f'{kind}:{_norm(name)}'
    cur = lib.get(key) or {}
    if cur.get('keyword') == kw:
        cur['hits'] = int(cur.get('hits', 0)) + 1
    else:
        cur = {'keyword': kw, 'hits': 1, 'last_text': str(text)[:40]}
    lib[key] = cur
    save()
    return kw


def suggest(kind, name):
    """查：这个科目以前学到过什么关键字。"""
    if not name:
        return None
    lib = load()
    ent = lib.get(f'{kind}:{_norm(name)}')
    return ent.get('keyword') if ent else None


def match_row(rows, kind, name):
    """按学到的关键字挑行；挑不到返回 None（调用方再去猜数值）。"""
    kw = suggest(kind, name)
    if not kw:
        return None
    for r in rows:
        if getattr(r, 'label', None) and kw in r.label:
            return r
    return None


def _norm(name):
    """科目名归一：去掉括号补充和空白，避免"实收资本（或股本）"对不上。"""
    s = re.sub(r'[（(][^）)]*[）)]', '', str(name or ''))
    return re.sub(r'\s+', '', s)
