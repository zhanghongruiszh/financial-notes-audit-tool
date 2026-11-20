# -*- coding: utf-8 -*-
"""
cross_checks.py — 附注合计数 与 三张报表科目数 的跨表勾稽。

做法
----
1. 每个子项目取其**第一张有"合计"行且列期间能定位**的表，读合计行的各期金额
2. 用子项目名去报表里按科目名找行（合并/母公司分别对应）
3. 按期间日期找报表的列
4. 差额 = 附注合计 − 报表数

报表找不到科目、或期间对应不上，不算差异，记为"未匹配"（待确认）。
"""

import datetime
import re
from dataclasses import dataclass
from typing import List, Optional

import report_reader as rr
import rules


@dataclass
class CrossCheck:
    chapter: str
    scope: str
    sub_item: str
    name: str
    period: Optional[datetime.date]
    period_key: Optional[str]
    note_value: Optional[float]
    note_ref: str                 # 如 "表#17 R4 C2"
    note_table: int = 0           # 附注表号（用于写到底稿右侧）
    note_row: int = 0             # 附注表内行号
    note_col: int = 0             # 附注表内列号
    note_terms: list = None       # [(行, 列, 系数)]：口径为"余额−准备"时要两项
    sheet: str = ''
    rows: Optional[int] = None
    cols: Optional[int] = None
    report_value: Optional[float] = None
    diff: Optional[float] = None
    status: str = 'unmatched'     # pass / fail / unmatched
    note: str = ''
    basis: str = ''               # 取数口径：账面价值 / 账面余额−坏账准备 / 账面余额


def _total_row_of(tm):
    """取代表总额的行：**优先「合计」，其次「小计」**。

    长期股权投资这类表里既有「小计」（分公司/联营各一段），也有「合计」，
    先拿小计会少算一半。
    """
    # 固定资产/无形资产这类表用「小计」而不是「合计」，
    # 并且有「账面原值合计 / 累计折旧合计 / 减值准备合计 / 账面价值合计」多段，
    # 对外披露的是**账面价值合计**。
    aggs = [r for r in tm.rows
            if r.kind in ('total', 'subtotal') and
            ('合计' in (r.label or '') or '总计' in (r.label or '')
             or '小计' in (r.label or ''))]
    if not aggs:
        return None
    for r in aggs:
        if '账面价值' in (r.label or ''):
            return r
    for r in aggs:
        label = (r.label or '')
        if ('合计' in label or '总计' in label) and '净值' not in label:
            return r
    for r in aggs:
        if '小计' not in (r.label or ''):
            return r
    return aggs[-1]


def _pick_main_table(tables, names=None):
    """挑一张最能代表该子项目金额的表。

    附注的排版习惯是：**子项目下的第一张表就是主披露表**（后面是分项、
    变动明细等）。所以按文档顺序取第一张合格的，而不是按列数打分——
    否则应收账款会拿到「组合计提」那种分表，金额口径就错了。
    """
    # 第零轮：有固化规则指定的代表行（租赁负债净额这类）
    for tm in tables:
        hinted = _hint_row(tm, names)
        if hinted is not None:
            return tm, hinted
    # 第一轮：有**「合计」**行且期间可定位（子项目下第一张这种表就是主表）
    for tm in tables:
        # 用统一的取行函数：它会优先挑「账面价值合计」，
        # 而不是「账面原值合计」那一段
        tot = _total_row_of(tm)
        if tot is None:
            continue
        if any(c.period for c in tm.columns if c.is_data):
            return tm, tot
    # 第一轮补：只有「小计」的也算
    for tm in tables:
        tot = _total_row_of(tm)
        if tot is None:
            continue
        if any(c.period for c in tm.columns if c.is_data):
            return tm, tot
    # 第二轮：没有合计行，但有「账面价值」/「期末xx」行或只有一个明细行
    for tm in tables:
        tot = _summary_row_of(tm, names)
        if tot is None:
            continue
        if any(c.period for c in tm.columns if c.is_data):
            return tm, tot
    return None, None


SUMMARY_LABELS = ('账面价值', '期末账面价值', '净额', '净值')

# 子项目名与报表科目名对不上时的别名
ALIAS_NAMES = {
    '现金流量表补充资料': ['经营活动产生的现金流量净额', '现金及现金等价物净增加额'],
    '营业收入和营业成本': ['营业收入', '营业成本'],
    # 附注把三项费用合在一节里披露
    '销售费用、管理费用、财务费用': ['销售费用', '管理费用', '研发费用', '财务费用'],
    # 报表用旧称、附注用新称的情况
    '应付职工薪酬': ['应付工资', '应付职工工资'],
    '应交税费': ['应交税金', '应交增值税'],
    '固定资产': ['固定资产原价', '固定资产净值'],
}


# 固化的取行规则：科目 -> 该科目在附注里"代表金额"那一行的写法
# （换项目也照样生效，不依赖某个具体文件）
ROW_HINTS = {
    '租赁负债': ['租赁负债净额', '租赁负债'],
    '递延所得税负债': ['递延所得税负债', '递延所得税负债净额'],
    '递延所得税资产': ['递延所得税资产', '递延所得税资产净额'],
    '长期股权投资': ['合计'],
    '营业收入': ['营业收入', '合计'],
    '营业成本': ['营业成本', '合计'],
    '销售费用': ['销售费用', '合计'],
    '管理费用': ['管理费用', '合计'],
    '研发费用': ['研发费用', '合计'],
    '财务费用': ['财务费用', '合计'],
    '实收资本(或股本)': ['实收资本', '合计'],
}


def _hint_row(tm, names):
    """按 ROW_HINTS 找代表行（行名包含提示词即可）。"""
    def has_money(r):
        return any(rules.cell_number(tm.ref.grid, r.row, c.index + 1)
                   for c in (tm.columns or []) if c.is_data)

    for n in (names or []):
        for want in ROW_HINTS.get(n, []):
            for r in tm.rows:
                if r.kind not in ('detail', 'subtotal', 'total'):
                    continue
                if want not in (r.label or ''):
                    continue
                # 名字对上了但整行是 0：这不是代表行，继续找
                if not has_money(r):
                    continue
                return r
    return None


def _summary_row_of(tm, names=None):
    """代表该子项目对外披露金额的行。

    依次尝试：
      1. 「合计」行
      2. 「账面价值」行（无形资产这类没有合计）
      3. 标签里同时含子项目名与「期末/年末」的行（如「期末未分配利润」）
      4. 表里**只有一个**明细行时，就用它（合同负债这类只披露一行）
    """
    hinted = _hint_row(tm, names)
    if hinted is not None:
        return hinted
    tot = _total_row_of(tm)
    if tot is not None:
        return tot
    for r in tm.rows:
        if r.kind == 'detail' and (r.label or '').strip() in SUMMARY_LABELS:
            return r
    if names:
        # 标签与子项目名（或别名）完全一致，或长名字被包含
        for r in tm.rows:
            if r.kind != 'detail':
                continue
            label = (r.label or '').replace(' ', '')
            for n in names:
                if not n:
                    continue
                if label == n or (len(n) >= 6 and n in label):
                    return r
        for r in tm.rows:
            if r.kind != 'detail':
                continue
            label = (r.label or '').replace(' ', '')
            # 必须以「期末/年末」开头：否则会选到「调整前上期末…」这种上期数
            if label.startswith(('期末', '年末')) and \
                    any(n and n in label for n in names):
                return r
    # 变动表版式：末行常写「本期期末余额」，取最后一个"期末"行
    tail = [r for r in tm.rows
            if r.kind == 'detail' and r.label
            and ('期末' in r.label or '年末' in r.label)
            and '期初' not in r.label and '年初' not in r.label]
    if tail:
        return tail[-1]

    details = [r for r in tm.rows if r.kind == 'detail' and r.label]
    if len(details) == 1:
        return details[0]
    return None


def _split_names(name):
    """把子项目名拆成能在报表里单独匹配的名字。

    含"和/、/及"时**不要**再拿全名去匹配——全名会被包含匹配误命中
    "营业收入"，导致取到成本列。直接按拆分名逐个匹配。
    """
    if name in ALIAS_NAMES:
        # 原名优先，别名兜底——否则报表用新称时反而匹配不上
        return [name] + ALIAS_NAMES[name]
    parts = [p.strip() for p in re.split(r'[和、及与]', name) if len(p.strip()) >= 2]
    if parts:
        return parts
    return [name]


def _match_report(report, scope, names):
    """在报表里按名字找科目，返回 (sheet, (row, col), 命中的名字)。"""
    for cand_name in names:
        for kind in ('资产负债表', '利润表', '现金流量表'):
            sheet = report.find(scope, kind)
            if sheet is None and scope != '单体':
                sheet = report.find('单体', kind)
            if sheet is None:
                continue
            pos = rr.find_item(sheet, cand_name)
            if pos:
                return sheet, pos, cand_name
    return None, None, None


def _note_col_of(tm, period):
    for c in tm.columns:
        if c.is_data and c.period == period:
            return c.index + 1
    return None


def _row_by_value(tm, current, sheet, item_row, item_col, get_value):
    """在表里挑一行，使它的金额最接近报表该项目的金额。

    不依赖行名（合计/净额/余额/账面价值…），纯按数值判断，
    换个项目、换个叫法都能用。
    """
    try:
        val_cols = (sheet.blocks or {}).get(item_col) or {}
    except Exception:
        return None
    if not val_cols:
        return None

    rows = [r for r in tm.rows
            if r.kind in ('detail', 'subtotal', 'total') and r.label]
    if not rows:
        return None

    best = None
    best_score = None
    for r in rows:
        total = 0.0
        hits = 0
        for col, period in val_cols.items():
            rv = get_value(sheet.name, item_row, col)
            if rv is None:
                continue
            nc = _note_col_of(tm, period)
            if nc is None:
                continue
            nv = rules.cell_number(tm.ref.grid, r.row, nc)
            if nv is None:
                continue
            total += abs(nv - rv)
            hits += 1
        if not hits:
            continue
        # 值相同的行有多条时（比如明细恰好等于合计），优先取"合计"行
        prefer = 0 if '合计' in (r.label or '') else 1
        sc = (total, prefer, -hits)
        if best_score is None or sc < best_score:
            best, best_score = r, sc
    if best is not None and best_score is not None and best_score[0] <= 0.01:
        return best
    return None


def build_number_index(model, report, wb):
    """用报表「附注」列建立 报表项目 → Word 小节 的编号索引。

    测试数据里的附注号是人工核对过的，可以直接当"对齐信号"：
    名称对不上时，用编号去定位它到底属于 Word 的哪一节。
    """
    import re
    word_by_no = {}
    for ch in model.chapters:
        for sub in ch.sub_items:
            no = getattr(sub, 'number', None)
            if no:
                word_by_no.setdefault(int(no), sub.name)

    index = {}
    for sheet in getattr(report, 'sheets', []):
        if sheet.name not in getattr(wb, 'sheetnames', []):
            continue
        ws = wb[sheet.name]
        for name, (row, icol) in sheet.items.items():
            for dc in (1, 2, -1, 3):
                if icol + dc < 1:
                    continue
                cell = ws.cell(row, icol + dc).value
                if not isinstance(cell, str) or len(cell) > 12:
                    continue
                m = re.search(r'(\d+)\s*$', cell)
                if not m:
                    continue
                no = int(m.group(1))
                if no in word_by_no:
                    index[(sheet.name, row)] = word_by_no[no]
                break
    return index


def fill_missing_checks(model, tables, report, out, get_value=None,
                        number_index=None):
    """给"还没挂上勾稽"的报表项目补勾稽。

    判定方式：在附注表里找一行，其金额与该报表项目相等（±0.01），
    并且行名/所在子项目名与项目名相关。这样一对多的科目
    （营业收入&营业成本、三项费用、财务费用的其中项）也能各自挂上。
    """
    if get_value is None:
        return out
    claimed = {(c.sheet, c.rows, c.cols) for c in out if c.sheet}
    try:
        import report_numbering as rn
    except Exception:
        rn = None

    for sheet in report.sheets:
        val_cols = {}
        for item_col, cols in (sheet.blocks or {}).items():
            for col, period in cols.items():
                val_cols.setdefault(period, col)
        if not val_cols:
            continue
        for name, (row, item_col) in sheet.items.items():
            if rn is not None and rn.is_no_note_item(name):
                continue
            if name == '其他综合收益':
                continue      # 确实没披露，不能拿 OCI 那一节顶替
            # 现金流量表只有"经营活动产生的现金流量净额"和附注勾，
            # 其余子项走 TB 核对，不在这里补配
            if sheet.name.startswith('现金流量表') and \
                    '经营活动产生的现金流量净额' not in name:
                continue
            cols = (sheet.blocks or {}).get(item_col) or {}
            if not cols:
                continue
            # 报表该项目各期的金额
            targets = {}
            for col, period in cols.items():
                v = get_value(sheet.name, row, col)
                if v:
                    targets[period] = v
            if not targets:
                continue
            if all((sheet.name, row, c) in claimed for c in cols):
                continue

            sub_by_no = (number_index or {}).get((sheet.name, row))
            for tm in tables:
                # 名称相关，或者编号指出它属于这一节
                rel = ((name in tm.sub_item)
                       or any(name in (r.label or '') for r in tm.rows)
                       or (sub_by_no and tm.sub_item == sub_by_no))
                if not rel:
                    continue
                for r in tm.rows:
                    if r.kind not in ('detail', 'subtotal', 'total') or not r.label:
                        continue
                    label = (r.label or '')
                    if name not in label and name not in tm.sub_item:
                        continue
                    for period, tval in targets.items():
                        nc = _note_col_of(tm, period)
                        if nc is None:
                            continue
                        nv = rules.cell_number(tm.ref.grid, r.row, nc)
                        if nv is None or abs(nv - tval) > 0.01:
                            continue
                        rcol = cols.get(None) or None
                        # 找到报表中该期间的列
                        rcol = next((c for c, pp in cols.items() if pp == period), None)
                        if rcol is None:
                            continue
                        key = (sheet.name, row, rcol)
                        if key in claimed:
                            continue
                        claimed.add(key)
                        out.append(CrossCheck(
                            chapter=tm.chapter, scope=sheet.scope if hasattr(sheet, 'scope') else '',
                            sub_item=tm.sub_item, name=name, period=period,
                            period_key=None, note_value=nv, note_ref='',
                            sheet=sheet.name, rows=row, cols=rcol,
                            note_table=tm.ref.idx, note_row=r.row, note_col=nc,
                            note_terms=[(r.row, nc, 1)], basis='补配'))
    return out


def _report_periods(report):
    """报表里出现过的期间，从大到小。

    用户给的规律：**大的那个必然是本期，小的是上期**，
    所以报表自己就能定出本期/上期，不依赖任何写死的年份。
    """
    periods = set()
    for sheet in getattr(report, 'sheets', []):
        for cols in (sheet.blocks or {}).values():
            for _c, p in cols.items():
                if p:
                    periods.add(p)
    return sorted(periods, reverse=True)


def _fallback_period(col, order):
    """附注表头没解析出日期时，按文字（本期/上期、期末/期初）挂到报表期间上。"""
    if col.period or not order:
        return None
    text = (col.header_text or '') + (col.period_key or '')
    if any(w in text for w in ('本期', '本年', '期末', '年末')):
        return order[0]
    if any(w in text for w in ('上期', '上年', '期初', '年初')):
        return order[1] if len(order) > 1 else None
    return None


def build_cross_checks(model, tables, report, get_value=None):
    by_sub = {}
    for tm in tables:
        by_sub.setdefault((tm.chapter, tm.sub_item), []).append(tm)

    out = []
    for ch in model.chapters:
        for sub in ch.sub_items:
            # 「归属于母公司所有者的…」这类是别的科目的组成部分，
            # 不是独立披露科目，不能拿它去和报表项目匹配
            if sub.name.startswith(('其中', '减：', '加：')):
                continue
            # 不是报表项目、也不是大标题的小节，不参与报表勾稽
            if any(w in sub.name for w in ('不重要合营', '不重要联营',
                                           '汇总信息', '不重要子公司',
                                           '受到限制', '受限资产')):
                continue
            is_oci_group = sub.name.startswith('归属于')
            group = by_sub.get((ch.name, sub.name), [])
            names = _split_names(sub.name)
            tm, tot = _pick_main_table(group, names)
            if tm is None:
                continue

            scope = ch.scope
            # 逐个候选名去报表里找：一个编号可能覆盖多个报表项目
            hits = []
            for nm in names:
                sh, fd, mn = _match_report(report, scope, [nm])
                if sh is not None and fd is not None:
                    if is_oci_group and mn == '其他综合收益':
                        continue
                    hits.append((sh, fd, mn))
            if not hits:
                sh, fd, mn = _match_report(report, scope, names)
                if sh is not None and fd is not None and not (
                        is_oci_group and mn == '其他综合收益'):
                    hits = [(sh, fd, mn)]
            sheet, found, matched_name = (hits[0] if hits
                                          else (None, None, None))
            if sheet is None or found is None:
                out.append(CrossCheck(
                    chapter=ch.name, scope=scope, sub_item=sub.name, name=sub.name,
                    period=None, period_key=None, note_value=None, note_ref='',
                    status='unmatched', note='报表里没找到同名科目'))
                continue
            item_row, item_col = found

            # 命中"营业收入"这类拆分名时，只取表头能对上的列。
            # 表头常写「收入」「成本」而不是全称，所以再试名字的后两个字。
            hint = matched_name if matched_name != sub.name else None
            hint_keys = []
            if hint:
                hint_keys = [hint]
                if len(hint) >= 2:
                    hint_keys.append(hint[-2:])

            def _col_ok(col):
                if not hint_keys:
                    return True
                head = (col.header_text or '').replace(' ', '')
                return any(k in head for k in hint_keys)

            # 模糊匹配代表行：名字对不上时，用**数值**去验证
            # （Word 里每个带编号的项目一定对应一个报表项目，两边金额应当相等）
            # 先用学到的模糊关键字挑（快且可解释）；
            # 挑不到再用数值猜，猜中了把语义沉淀成规则
            try:
                import learn_rules
                learned = learn_rules.match_row(
                    [r for r in tm.rows
                     if r.kind in ('detail', 'subtotal', 'total') and r.label],
                    'row', matched_name or sub.name)
                if learned is not None:
                    tot = learned
            except Exception:
                pass
            if get_value is not None:
                guessed = _row_by_value(tm, tot, sheet, item_row, item_col,
                                        get_value)
                if guessed is not None:
                    if guessed is not tot:
                        try:
                            import learn_rules
                            learn_rules.learn('row',
                                              matched_name or sub.name,
                                              guessed.label)
                        except Exception:
                            pass
                    tot = guessed

            order = _report_periods(report)
            per_period = {}
            for col in tm.columns:
                if not col.is_data:
                    continue
                if not col.period:
                    col.period = _fallback_period(col, order)
                if not col.period:
                    continue
                if not _col_ok(col):
                    continue
                role = col.role or ('金额' if col.kind == 'amount' else None)
                if role in ('价值', '余额', '金额', '准备'):
                    # 同一期间同一角色有多列时（如"收入/成本"都判成金额），
                    # 保留文档顺序里的第一列，别让后面的覆盖前面的
                    per_period.setdefault(col.period, {}).setdefault(role, col)
            if hint and not per_period:
                # 表头完全对不上，退回用全部列，但标记出来
                for col in tm.columns:
                    if not col.is_data or not col.period:
                        continue
                    role = col.role or ('金额' if col.kind == 'amount' else None)
                    if role in ('价值', '余额', '金额', '准备'):
                        per_period.setdefault(col.period, {})[role] = col

            for period, roles in sorted(per_period.items()):
                if '价值' in roles:
                    col, basis = roles['价值'], '账面价值'
                    note_val = rules.cell_number(tm.ref.grid, tot.row, col.index + 1)
                    terms = [(tot.row, col.index + 1, 1)]
                elif '余额' in roles and '准备' in roles:
                    col, basis = roles['余额'], '账面余额−坏账准备'
                    bal = rules.cell_number(tm.ref.grid, tot.row, roles['余额'].index + 1)
                    prov = rules.cell_number(tm.ref.grid, tot.row, roles['准备'].index + 1)
                    note_val = None if (bal is None or prov is None) else bal - prov
                    terms = [(tot.row, roles['余额'].index + 1, 1),
                             (tot.row, roles['准备'].index + 1, -1)]
                elif '余额' in roles:
                    col, basis = roles['余额'], '账面余额'
                    note_val = rules.cell_number(tm.ref.grid, tot.row, col.index + 1)
                    terms = [(tot.row, col.index + 1, 1)]
                elif '金额' in roles:
                    col, basis = roles['金额'], '金额'
                    note_val = rules.cell_number(tm.ref.grid, tot.row, col.index + 1)
                    terms = [(tot.row, col.index + 1, 1)]
                else:
                    continue
                if note_val is None:
                    continue

                cc = CrossCheck(
                    chapter=ch.name, scope=scope, sub_item=sub.name,
                    name=sub.name, period=col.period, period_key=col.period_key,
                    note_value=note_val, basis=basis,
                    note_ref=f'表#{tm.ref.idx} R{tot.row} C{col.index + 1}',
                    note_table=tm.ref.idx, note_row=tot.row, note_col=col.index + 1,
                    note_terms=terms,
                )
                col_no = sheet.value_col(item_col, col.period)
                if col_no is None:
                    cc.sheet = sheet.name
                    cc.rows = item_row
                    cc.cols = item_col
                    cc.status = 'unmatched'
                    cc.note = f'报表里没有 {col.period} 这一期'
                    out.append(cc)
                    continue

                cc.sheet = sheet.name
                cc.rows = item_row
                cc.cols = col_no
                out.append(cc)
    return out


NIL_VALUES = {'-', '--', '—', '——', '–', '/', '', '无', '不适用'}


def to_number(value):
    """把报表单元格转成数字。'-' 等空值标记按 0 处理，无法解析返回 None。"""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if value is None:
        return 0.0
    s = str(value).strip().replace(',', '')
    if s in NIL_VALUES:
        return 0.0
    if s.endswith('%'):
        try:
            return float(s[:-1]) / 100.0
        except ValueError:
            return None
    try:
        return float(s)
    except ValueError:
        return None


def evaluate(cross, report_values):
    """report_values: {(sheet_name,row,col): value}"""
    if cross.sheet and cross.rows and cross.cols:
        raw = report_values.get((cross.sheet, cross.rows, cross.cols))
        cross.report_value = to_number(raw)
        if cross.report_value is None:
            cross.status = 'unmatched'
            cross.note = f'报表取数不是数字：{raw!r}'
            return cross
    else:
        cross.report_value = None
    if cross.report_value is None:
        if cross.status != 'unmatched':
            cross.status = 'unmatched'
            cross.note = cross.note or '报表取数为空'
        return cross
    cross.diff = cross.note_value - cross.report_value
    cross.status = 'pass' if abs(cross.diff) <= rules.TOLERANCE else 'fail'
    return cross
