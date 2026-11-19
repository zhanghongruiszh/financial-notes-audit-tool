# -*- coding: utf-8 -*-
"""
profile_tables.py — 把模板里的表结构"提纯"成骨架，用于归纳加减关系。

做法：
  1. 遍历模板，解析每张附注表
  2. 给每张表生成一个结构签名：
       行签名 = 行标签序列（数字/日期/公司名统一替换成占位符）
       列签名 = 列路径序列
  3. 按签名聚类，统计出现频次与覆盖的模板数
  4. 输出 _work/table_families.txt（供人看）与 .csv（供程序用）

用法：
    python tools/profile_tables.py               # 用 模板库/ 里的骨架 + 源 docx
    python tools/profile_tables.py --limit 200   # 只看前 200 个家族
"""

import os
import re
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import notes_parser            # noqa: E402
import table_semantics as ts   # noqa: E402
import template_lib            # noqa: E402

RE_DIGIT = re.compile(r'[\d,，.]+')
RE_DATE = re.compile(r'\d{4}\s*年|\d{4}-\d{2}-\d{2}|\d{1,2}\s*月')
RE_COMPANY = re.compile(r'[\u4e00-\u9fa5（）()]{2,20}(有限公司|股份公司|集团|公司)$')

# 结构关键行/列的词表：把这些词保留，其余抽象成占位
KEY_WORDS = ('期初', '期末', '本期增加', '本期减少', '合计', '小计', '总计',
             '其中', '账面原值', '累计折旧', '累计摊销', '减值准备', '账面价值',
             '账面余额', '坏账准备', '计提', '收回', '转回', '转销', '核销',
             '增加', '减少', '余额', '金额', '比例')


def norm_label(text):
    """把具体名称抽象掉，只留结构骨架。"""
    s = (text or '').strip()
    if not s:
        return ''
    s = RE_DATE.sub('<日期>', s)
    s = RE_COMPANY.sub('<主体>', s)
    s = RE_DIGIT.sub('<数>', s)
    s = re.sub(r'[①②③④⑤⑥⑦⑧⑨⑩]', '<块>', s)
    for w in KEY_WORDS:
        if w in s:
            return w
    return '<其他>'


def col_signature(tm):
    sig = []
    for c in tm.columns:
        if not c.is_data:
            continue
        head = norm_label(c.header_text) or '<空>'
        sig.append(f'{head}/{c.role or "-"}')
    return '|'.join(sig)


def row_signature(tm):
    sig = []
    for r in tm.rows:
        if r.kind in ('title', 'header'):
            continue
        tag = norm_label(r.label)
        if not tag:
            continue
        sig.append(f'{tag}:{r.kind[:3]}')
    return '|'.join(sig)


def main():
    limit = 200
    if '--limit' in sys.argv:
        limit = int(sys.argv[sys.argv.index('--limit') + 1])
    coverage = '--coverage' in sys.argv

    tpl_dir = '_work/docx'
    if not os.path.isdir(tpl_dir):
        print('找不到 _work/docx（模板转换目录）')
        return 2

    import rules
    from collections import Counter
    table_rule_counter = Counter()      # 按"表"统计，避免家族并集重复计数

    docs = [f for f in sorted(os.listdir(tpl_dir))
            if f.endswith('.docx') and not f.startswith('3、')]
    families = defaultdict(lambda: {'count': 0, 'docs': set(), 'examples': [],
                                    'types': set(), 'rules': set()})
    total = 0

    for doc in docs:
        path = os.path.join(tpl_dir, doc)
        try:
            model = notes_parser.parse_notes(path, work_dir='_work/profile')
            tables = ts.analyze_document(model)
        except Exception as exc:
            print(f'跳过 {doc}: {exc}')
            continue
        print(f'{doc}: {len(tables)} 张表')
        for t in tables:
            total += 1
            if coverage:
                # 模板是空白表，列被当成文本；这里假设"有数"以便看规则是否适用
                for c in t.columns:
                    if c.is_data and c.kind == 'text':
                        c.kind = 'amount'
            key = (row_signature(t), col_signature(t))
            fam = families[key]
            fam['count'] += 1
            fam['docs'].add(doc)
            fam['types'].add(t.table_type_name)
            if len(fam['examples']) < 3:
                fam['examples'].append(f'{doc}#{t.ref.idx} {t.sub_item}')
            if coverage:
                # 空壳表（没有任何明细行）本来就无可核，单独统计
                has_rows = any(r.kind in ('detail', 'subtotal')
                               and r.label for r in t.rows)
                if not has_rows:
                    fam['rules'].add('(空壳)')
                    table_rule_counter['(空壳)'] += 1
                else:
                    try:
                        fired = {chk.rule_id for chk in rules.build_checks([t])}
                        fam['rules'] |= fired or {'(无)'}
                        if fired:
                            for rid in fired:
                                table_rule_counter[rid] += 1
                        else:
                            table_rule_counter['(无)'] += 1
                    except Exception:
                        fam['rules'].add('ERR')
                        table_rule_counter['ERR'] += 1

    print(f'\n共 {total} 张表，{len(families)} 个结构家族\n')

    ranked = sorted(families.items(), key=lambda kv: (-len(kv[1]['docs']), -kv[1]['count']))
    out_lines = []
    if coverage:
        # 覆盖率汇总
        covered = sum(1 for _, f in families.items() if f['rules'] - {'(空壳)', '(无)', 'ERR'})
        blank = sum(1 for _, f in families.items() if not f['rules'])
        out_lines.append(f'== 覆盖度：{covered} 个家族有规则命中，{blank} 个没有 ==')
        out_lines.append('== 按规则覆盖的表数（按表统计）==')
        for r, n in table_rule_counter.most_common():
            out_lines.append(f'  {r:<10} {n:>5} 张')
        out_lines.append('')
        # 只看没覆盖的
        ranked = [(k, f) for k, f in ranked if not (f['rules'] - {'(空壳)', '(无)'})]

    for i, (key, fam) in enumerate(ranked[:limit], 1):
        row_sig, col_sig = key
        rules_txt = ('　规则：' + '/'.join(sorted(fam['rules']))) if coverage else ''
        head = (f'[家族 {i}] 出现 {fam["count"]} 次，覆盖 {len(fam["docs"])} 份模板　'
                f'类型：{"/".join(sorted(fam["types"]))}{rules_txt}')
        out_lines.append(head)
        out_lines.append(f'  行：{row_sig[:300]}')
        out_lines.append(f'  列：{col_sig[:300]}')
        out_lines.append(f'  例：{"；".join(fam["examples"])}')
        out_lines.append('')

    text = '\n'.join(out_lines)
    with open('_work/table_families.txt', 'w', encoding='utf-8') as f:
        f.write(text)

    # 同时写一份可入库的说明文档
    os.makedirs('docs', exist_ok=True)
    md = ['# 模板表结构家族（骨架提纯结果）', '',
          '由 `python tools/profile_tables.py --coverage` 生成。',
          '把 5 份模板里所有附注表按「行标签序列 + 列结构」聚类，得到结构家族，',
          '再用当前规则跑一遍，看每个家族能触发哪些勾稽。', '']
    if coverage:
        md.append('## 覆盖度')
        md.append('')
        md.append('| 情况 | 表数 |')
        md.append('|---|---|')
        for r, n in table_rule_counter.most_common():
            label = {'(无)': '有明细行但规则未覆盖',
                     '(空壳)': '空壳（模板没填数、只有合计行）'}.get(r, f'规则 `{r}`')
            md.append(f'| {label} | {n} |')
        md.append('')
    md.append('## 结构家族清单')
    md.append('')
    md.append('```')
    md.append(text[:60000])
    md.append('```')
    with open('docs/表结构家族.md', 'w', encoding='utf-8') as f:
        f.write('\n'.join(md))

    print(text[:8000])
    print('\n完整清单已写入 _work/table_families.txt 与 docs/表结构家族.md')
    return 0


if __name__ == '__main__':
    sys.exit(main())
