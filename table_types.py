# -*- coding: utf-8 -*-
"""
table_types.py — 识别一张附注表属于哪一类。

为什么需要
----------
同样是"8 行 8 列"，可能是账龄表、坏账计提表、还是费用明细表——类型不同，
该看的勾稽关系也不同。标出类型有三个用处：

  1. 成果表里一眼看出"这是哪类表"，人工核对时不用再猜
  2. 总览里给出类型分布，缺哪类披露一目了然
  3. 给规则引擎一个结构性依据（例如"资产变动表"必然有 期初/增加/减少/期末）

识别依据
--------
· 行标签用词（账龄、单项、组合、期初、本期增加、第一阶段…）
· 列角色（余额/准备/价值/比例/增加/减少）
· 表头文字（债务人名称、预期信用损失…）
· 子项目名（管理费用、财务费用…）

这些都是"通用财务披露常识"，不依赖具体客户。
"""

import re

TYPE_NAMES = {
    'CF_SUPPLEMENT': '现金流量表补充资料',
    'STAGED_PROVISION': '三阶段坏账变动表',
    'AGING': '账龄分析表',
    'PROVISION': '坏账/减值计提表',
    'MOVEMENT': '资产增减变动表',
    'TOP5': '前五名/关联方明细表',
    'TAX': '所得税调整表',
    'EXPENSE': '费用明细表',
    'BALANCE': '余额明细表',
    'OTHER': '未分类',
}

RE_AGING = re.compile(r'年以内|年以内（|至\d+年|\d+年以上|^\d+年')
RE_TOP5_WORDS = ('有限公司', '股份', '集团', '公司')


def _row_labels(tm):
    return [r.label for r in tm.rows
            if r.kind in ('detail', 'group', 'subtotal') and r.label]


def _roles(tm):
    return {c.role for c in tm.columns if c.is_data and c.role}


def _headers(tm):
    return ' '.join((c.header_text or '') for c in tm.columns)


def classify(tm, sub_item=''):
    """返回 (类型代码, 类型中文名, 置信度 0~1)。"""
    labels = _row_labels(tm)
    joined = ' '.join(labels)
    roles = _roles(tm)
    headers = _headers(tm)
    total_rows = [r for r in tm.rows if r.kind == 'total']

    def hit(*words):
        return sum(1 for w in words if w in joined)

    # 1. 现金流量表补充资料
    if '现金流量表' in (sub_item or '') or \
            any(w in joined for w in ('将净利润调节', '经营活动产生的现金流量净额')):
        return 'CF_SUPPLEMENT', TYPE_NAMES['CF_SUPPLEMENT'], 0.95

    # 2. 三阶段预期信用损失变动表
    if '第一阶段' in headers or '第二阶段' in headers or '预期信用损失' in headers:
        return 'STAGED_PROVISION', TYPE_NAMES['STAGED_PROVISION'], 0.9

    # 3. 账龄分析表
    if '账龄' in headers or sum(1 for l in labels if RE_AGING.search(l)) >= 2:
        return 'AGING', TYPE_NAMES['AGING'], 0.9

    # 4. 资产增减变动表（原值/折旧/减值，或 期初+增加−减少=期末）
    if ('期初' in joined and ('本期增加' in joined or '本期减少' in joined)) or \
            ('账面原值' in joined or '累计折旧' in joined or '累计摊销' in joined):
        return 'MOVEMENT', TYPE_NAMES['MOVEMENT'], 0.9

    # 5. 前五名 / 关联方明细（特征很强：表头写债务人名称，或行大多是公司名）
    #    必须排在「计提表」之前——前五名表也常有"坏账准备 + 占比"两列，
    #    否则会被误判成计提表。
    if '债务人名称' in headers or '单位名称' in headers or \
            sum(1 for l in labels if any(w in l for w in RE_TOP5_WORDS)) >= 3:
        return 'TOP5', TYPE_NAMES['TOP5'], 0.8

    # 6. 坏账/减值计提表（余额/金额 + 比例 + 准备）
    if '准备' in roles and '比例' in roles:
        return 'PROVISION', TYPE_NAMES['PROVISION'], 0.85
    if '准备' in roles and '价值' in roles and hit('单项', '组合', '计提') >= 1:
        return 'PROVISION', TYPE_NAMES['PROVISION'], 0.8

    # 7. 所得税调整表
    if '适用税率' in joined or ('利润总额' in joined and '调整' in joined):
        return 'TAX', TYPE_NAMES['TAX'], 0.85

    # 8. 费用明细表
    if any(w in (sub_item or '') for w in ('管理费用', '销售费用', '财务费用',
                                           '研发费用', '税金及附加', '费用')):
        return 'EXPENSE', TYPE_NAMES['EXPENSE'], 0.75

    # 9. 通用余额明细表：有合计行、有数据列，就是一张普通的"项目×期间"表
    if total_rows and any(c.is_data for c in tm.columns):
        return 'BALANCE', TYPE_NAMES['BALANCE'], 0.6
    if any(c.is_data for c in tm.columns) and labels:
        return 'BALANCE', TYPE_NAMES['BALANCE'], 0.45

    return 'OTHER', TYPE_NAMES['OTHER'], 0.2


def summarize(tables):
    """按类型统计表数。"""
    from collections import Counter
    c = Counter(t.table_type for t in tables if getattr(t, 'table_type', None))
    return {TYPE_NAMES.get(k, k): v for k, v in c.most_common()}
