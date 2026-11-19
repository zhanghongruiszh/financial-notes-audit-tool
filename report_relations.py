# -*- coding: utf-8 -*-
"""
report_relations.py — 三张报表内部的会计恒等式（不依赖附注）。

这一层是**最底层的可信校验**：只要报表的勾稽关系对不上，后面的附注勾稽
都没有意义。全部用声明式关系表描述，按"模糊项目名"取值，不写死行号列号。

集成的关系（会计上必须成立的）：

资产负债表
  · 流动资产合计   = 各流动资产项目之和
  · 非流动资产合计 = 各非流动资产项目之和
  · 资产总计       = 流动资产合计 + 非流动资产合计
  · 流动负债合计   = 各流动负债项目之和
  · 非流动负债合计 = 各非流动负债项目之和
  · 负债合计       = 流动负债合计 + 非流动负债合计
  · 所有者权益合计 = 各权益项目之和
  · 负债和所有者权益总计 = 负债合计 + 所有者权益合计
  · 资产总计       = 负债和所有者权益总计          ← 会计恒等式

利润表
  · 营业利润   = 营业总收入 − 营业总成本 + 其他收益 + 投资收益
                 + 公允价值变动收益 + 信用减值损失 + 资产减值损失 + 资产处置收益
  · 利润总额   = 营业利润 + 营业外收入 − 营业外支出
  · 净利润     = 利润总额 − 所得税费用
  · 综合收益总额 = 净利润 + 其他综合收益的税后净额

现金流量表
  · 现金及现金等价物净增加额 = 经营活动净额 + 投资活动净额 + 筹资活动净额
  · 期末现金及现金等价物余额 = 期初现金及现金等价物余额 + 净增加额
"""

import re

TOL = 0.01

# (报表, 目标项目, [(项目名关键字, 系数), ...], 说明)
RELATIONS = [
    ('资产负债表', '资产总计',
     [('流动资产合计', +1), ('非流动资产合计', +1)],
     '资产总计 = 流动资产合计 + 非流动资产合计'),
    ('资产负债表', '负债合计',
     [('流动负债合计', +1), ('非流动负债合计', +1)],
     '负债合计 = 流动负债合计 + 非流动负债合计'),
    ('资产负债表', '负债和所有者权益(或股东权益)总计',
     [('负债合计', +1), ('所有者权益(或股东权益)合计', +1)],
     '负债和所有者权益总计 = 负债合计 + 所有者权益合计'),
    ('资产负债表', '资产总计',
     [('负债和所有者权益(或股东权益)总计', +1)],
     '会计恒等式：资产总计 = 负债和所有者权益总计'),

    ('利润表', '营业利润',
     [('营业总收入', +1), ('营业总成本', -1), ('其他收益', +1),
      ('投资收益', +1), ('公允价值变动收益', +1), ('信用减值损失', +1),
      ('资产减值损失', +1), ('资产处置收益', +1)],
     '营业利润 = 营业总收入 − 营业总成本 + 其他收益 + 投资收益 + 公允价值变动 + 减值 + 处置收益'),
    ('利润表', '营业利润',
     [('营业收入', +1), ('营业成本', -1), ('税金及附加', -1),
      ('销售费用', -1), ('管理费用', -1), ('研发费用', -1), ('财务费用', -1),
      ('其他收益', +1), ('投资收益', +1), ('公允价值变动收益', +1),
      ('信用减值损失', +1), ('资产减值损失', +1), ('资产处置收益', +1)],
     '营业利润 = 营业收入 − 营业成本 − 税金及附加 − 期间费用 + 其他收益等'),
    ('利润表', '利润总额',
     [('营业利润', +1), ('营业外收入', +1), ('营业外支出', -1)],
     '利润总额 = 营业利润 + 营业外收入 − 营业外支出'),
    ('利润表', '净利润',
     [('利润总额', +1), ('所得税费用', -1)],
     '净利润 = 利润总额 − 所得税费用'),
    ('利润表', '综合收益总额',
     [('净利润', +1), ('其他综合收益的税后净额', +1)],
     '综合收益总额 = 净利润 + 其他综合收益的税后净额'),

    ('现金流量表', '现金及现金等价物净增加额',
     [('经营活动产生的现金流量净额', +1),
      ('投资活动产生的现金流量净额', +1),
      ('筹资活动产生的现金流量净额', +1),
      ('*汇率变动对现金及现金等价物的影响', +1)],
     '净增加额 = 经营 + 投资 + 筹资 净额（+ 汇率变动影响）'),
    ('现金流量表', '期末现金及现金等价物余额',
     [('期初现金及现金等价物余额', +1), ('现金及现金等价物净增加额', +1)],
     '期末余额 = 期初余额 + 净增加额'),
]


def _norm(text):
    s = re.sub(r'[（(][^）)]*[）)]', '', str(text or ''))
    return re.sub(r'\s+', '', s)


def find_item(sheet, key):
    """在报表里按模糊关键字找项目，返回 (项目名, 行, 项目列)。"""
    nk = _norm(key)
    best = None
    for name, pos in sheet.items.items():
        nn = _norm(name)
        if nn == nk:
            return name, pos[0], pos[1]
        if nk in nn and best is None:
            best = (name, pos[0], pos[1])
    return best if best else (None, None, None)


def value_of(book, sheet, row, item_col, period):
    """取某项目在某期间的值。"""
    try:
        col = next((c for c, p in (sheet.blocks.get(item_col) or {}).items()
                    if p == period), None)
    except Exception:
        col = None
    if col is None or sheet.name not in book.sheetnames:
        return None
    v = book[sheet.name].cell(row, col).value
    return float(v) if isinstance(v, (int, float)) else None


def check(book, report, relations=None):
    """返回 [(报表, 说明, 期间, 期望, 实际, 差额, 状态)]。"""
    out = []
    rels = relations or RELATIONS
    sheets = {s.name: s for s in report.sheets}
    for sheet_name, target_key, terms, desc in rels:
        sheet = sheets.get(sheet_name)
        if sheet is None:
            continue
        tname, trow, tcol = find_item(sheet, target_key)
        if trow is None:
            out.append((sheet_name, desc, None, None, None, None, 'unmatched',
                        f'没找到「{target_key}」'))
            continue
        periods = sorted({p for c, p in (sheet.blocks.get(tcol) or {}).items()},
                         reverse=True)
        for period in periods:
            target = value_of(book, sheet, trow, tcol, period)
            if target is None:
                continue
            total = 0.0
            got = 0
            missing = []
            hard_missing = []
            for key, coef in terms:
                optional = key.startswith('*')
                k = key[1:] if optional else key
                nm, row, col = find_item(sheet, k)
                if row is None:
                    if not optional:
                        hard_missing.append(k)
                    continue
                v = value_of(book, sheet, row, col, period)
                if v is None:
                    v = 0.0
                total += coef * v
                got += 1
            if hard_missing:
                out.append((sheet_name, desc, period, None, target, None,
                            'skip', '缺：' + '、'.join(hard_missing)))
                continue
            if not got:
                continue
            diff = target - total
            status = 'pass' if abs(diff) <= TOL else 'fail'
            note = ('缺：' + '、'.join(missing)) if missing else ''
            out.append((sheet_name, desc, period, total, target, diff,
                        status, note))
    return out


def summarize(results):
    from collections import Counter
    return dict(Counter(r[6] for r in results))

def child_items(wb, report):
    """挑出"不是报表项目、只是子项"的行。

    判据（结合会计恒等式一起用）：
      1. 行名以「其中/减：/加：」开头 —— 明确的子项标记
      2. 参与恒等式的项目一定不是子项（如"营业总收入""净利润"）
    返回 {项目名}。
    """
    involved = set()
    for _sheet, target, terms, _desc in RELATIONS:
        involved.add(_norm(target))
        for key, _coef in terms:
            involved.add(_norm(key.lstrip('*')))

    out = set()
    for sheet in getattr(report, 'sheets', []):
        if sheet.name not in getattr(wb, 'sheetnames', []):
            continue
        ws = wb[sheet.name]
        for name, (row, icol) in sheet.items.items():
            if _norm(name) in involved:
                continue
            try:
                raw = ws.cell(row, icol).value
            except Exception:
                raw = None
            text = str(raw if raw is not None else name).strip()
            if text.startswith(('其中', '减：', '减:', '加：', '加:')):
                out.add(name)
    return out
