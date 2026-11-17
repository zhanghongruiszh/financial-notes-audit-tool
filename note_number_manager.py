# -*- coding: utf-8 -*-
"""
附注编号智能管理器 (Note Number Manager) - 高性能版
功能：
1. 极速提取与检查 Word 附注子项目编号连续性（断号、重号、乱序、未编号项目识别）
2. Word 附注子项目一键顺序重编号（严格限定在主要项目附注与母公司附注模块）
3. 报表 Excel 附注编号一键修复与补全（资产负债表、利润表、现金流量表等）
"""

import os
import re
import shutil
import openpyxl

# 标准常用会计科目关键词（用于语义锚点，即使没有数字编号也能精准识别）
ACCOUNTING_SUBJECTS = [
    # 资产类
    '货币资金', '结算备付金', '拆出资金', '交易性金融资产', '衍生金融资产', '应收票据',
    '应收账款', '应收款项融资', '预付款项', '应收保费', '应收分保账款', '应收分保合同准备金',
    '其他应收款', '买入返售金融资产', '存货', '合同资产', '持有待售资产', '一年内到期的非流动资产',
    '其他流动资产', '债权投资', '其他债权投资', '长期应收款', '长期股权投资', '其他权益工具投资',
    '其他非流动金融资产', '投资性房地产', '固定资产', '在建工程', '工程物资', '固定资产清理',
    '生产性生物资产', '油气资产', '使用权资产', '无形资产', '开发支出', '商誉', '长期待摊费用',
    '递延所得税资产', '其他非流动资产',
    # 负债类
    '短期借款', '向中央银行借款', '拆入资金', '交易性金融负债', '衍生金融负债', '应付票据',
    '应付账款', '预收款项', '合同负债', '卖出回购金融资产款', '吸收存款及同业存放', '代理买卖证券款',
    '代理承销证券款', '应付职工薪酬', '应交税费', '其他应付款', '应付手续费及佣金', '应付分保账款',
    '持有待售负债', '一年内到期的非流动负债', '其他流动负债', '长期借款', '应付债券', '租赁负债',
    '长期应付款', '长期应付职工薪酬', '预计负债', '递延收益', '递延所得税负债', '其他非流动负债',
    # 权益类
    '实收资本', '股本', '其他权益工具', '资本公积', '减：库存股', '其他综合收益', '专项储备',
    '盈余公积', '一般风险准备', '未分配利润',
    # 损益类
    '营业收入', '营业收入和营业成本', '营业成本', '税金及附加', '销售费用', '管理费用',
    '研发费用', '财务费用', '其他收益', '投资收益', '净敞口套期收益', '公允价值变动收益',
    '信用减值损失', '资产减值损失', '资产处置收益', '营业外收入', '营业外支出', '所得税费用',
    # 现金流与补充资料
    '现金流量表补充资料', '现金及现金等价物'
]

def clean_subject_name(name):
    """去除科目名称中的标点、空格、数字、'其中：'等"""
    if not name:
        return ''
    s = re.sub(r'^[（(]?[一二三四五六七八九十0-9]+[）)]?[、.．\s]*', '', name)
    s = re.sub(r'[\s\u3000\t\r\n\x07\x0c]', '', s)
    s = re.sub(r'^(其中[:：]|加[:：]|减[:：])', '', s)
    return s

def extract_word_note_items(doc_path):
    """
    通过 doc.Content.Text 极速秒级解析 Word 附注主要项目
    """
    import win32com.client as win32
    
    word = win32.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = False
    
    try:
        doc = word.Documents.Open(os.path.abspath(doc_path), ReadOnly=True)
        raw_text = doc.Content.Text
        doc.Close(False)
    finally:
        word.Quit()
        
    lines = [l.strip() for l in raw_text.splitlines() if l.strip()]
    
    chapter_prefix = ''
    parent_prefix = ''
    consolidated_items = []
    parent_items = []
    
    current_mode = None # 'consolidated' or 'parent'
    
    for idx, t in enumerate(lines):
        # 排除表格单元格内的文字（Word中\x07表示表格单元格结束符）
        if '\x07' in t or '\x0c' in t:
            continue
        # 检测主要项目附注章节入口
        if ('主要项目附注' in t or '主要项目注释' in t or '财务报表项目注释' in t) and any(t.startswith(p) for p in ['一、', '二、', '三、', '四、', '五、', '六、', '七、', '八、', '九、', '十、', '十一、', '十二、', '十三、', '十四、', '十五、', '十六、', '十七、', '十八、', '十九、', '二十、']):
            m_chap = re.match(r'^([一二三四五六七八九十百]+、)', t)
            pref = m_chap.group(1) if m_chap else ''
            if '母公司' in t:
                current_mode = 'parent'
                parent_prefix = pref
            else:
                current_mode = 'consolidated'
                chapter_prefix = pref
            continue
            
        # 检测退出或切换大章节
        if current_mode and any(t.startswith(prefix) for prefix in ['七、', '八、', '九、', '十、', '十一、', '十二、', '十三、', '十四、', '十五、', '十六、', '十七、', '十八、', '十九、', '二十、']):
            if '母公司' in t and ('主要项目' in t or '附注' in t or '注释' in t):
                current_mode = 'parent'
                m_chap = re.match(r'^([一二三四五六七八九十百]+、)', t)
                parent_prefix = m_chap.group(1) if m_chap else ''
                continue
            else:
                current_mode = None
                continue
                
        if current_mode:
            # 模式 1：以阿拉伯数字开头，例如 "1、货币资金", "1. 货币资金", "1 货币资金"
            m_num = re.match(r'^(\d+)[\s、.．\t]+([^\d\s、.．].*)$', t)
            if m_num:
                num_val = int(m_num.group(1))
                raw_subj = m_num.group(2).strip()
                # 排除二级明细小标题，如 "(1) 账龄分析", "1) 组合计提"
                if not raw_subj.startswith(('(', '（', '1)', '1）', '①', '②', '其中', '注')):
                    c_name = clean_subject_name(raw_subj)
                    # 验证包含已知科目
                    is_subj = any(s in c_name for s in ACCOUNTING_SUBJECTS) or len(c_name) <= 12
                    if is_subj:
                        item_tuple = (idx, num_val, c_name, t)
                        if current_mode == 'consolidated':
                            consolidated_items.append(item_tuple)
                        else:
                            parent_items.append(item_tuple)
                        continue
                        
            # 模式 2：未编号项目，通过严格会计科目词典识别
            # 严格过滤二级编号：如 (1), (2), （1）, （2）, ①, ②, 1), 2), [1], 注1 等
            if re.match(r'^[（(\[【]?[0-9一二三四①②③④⑤⑥⑦⑧⑨⑩]+[）)\]】、.．]', t):
                continue
                
            clean_t = clean_subject_name(t)
            is_sub_detail = any(sub in t for sub in [
                '账龄', '计提', '组合', '前五名', '明细', '分类', '终止确认', '转移', '收益以',
                '减少以', '增加以', '坏账准备', '存货跌价', '减值准备', '折旧', '摊销', '变动',
                '余额', '净额', '账面价值', '账面余额', '其他说明', '按性质', '按款项', '项目'
            ])
            if not is_sub_detail:
                for subj in ACCOUNTING_SUBJECTS:
                    if clean_t == subj:
                        # 如果当前科目与刚提取的最后一个科目相同（如"其他应收款"内部的"(3)其他应收款"），说明是内部子节，跳过
                        last_items = consolidated_items if current_mode == 'consolidated' else parent_items
                        if last_items and last_items[-1][2] == subj:
                            break
                        item_tuple = (idx, None, subj, t)
                        if current_mode == 'consolidated':
                            consolidated_items.append(item_tuple)
                        else:
                            parent_items.append(item_tuple)
                        break
                            
    return chapter_prefix, parent_prefix, consolidated_items, parent_items

def diagnose_note_continuity(items, section_name="合并财务报表主要项目附注"):
    """
    诊断子项目编号连续性
    返回: (is_continuous, issues_list)
    """
    issues = []
    if not items:
        return True, ["未提取到该章节子项目"]
        
    expected = 1
    seen_nums = set()
    
    for idx, (p_i, num, name, raw_text) in enumerate(items):
        if num is None:
            prev_name = items[idx-1][2] if idx > 0 else "章节开始"
            next_name = items[idx+1][2] if idx < len(items)-1 else "章节结束"
            issues.append(f"【缺失编号】科目【{name}】未填写序号！位于【{prev_name}】与【{next_name}】之间")
        else:
            if num in seen_nums:
                issues.append(f"【重复编号】序号【{num}】重复出现！科目：【{name}】")
            elif num != expected:
                issues.append(f"【跳号/断号】序号不连续！期望序号为【{expected}】，实际为【{num}】（科目：{name}）")
                expected = num
            seen_nums.add(num)
            expected += 1
            
    is_continuous = (len(issues) == 0)
    return is_continuous, issues

def renumber_word_file(doc_path, output_path=None):
    """
    对 Word 文档附注主要项目进行极速顺序重编号
    """
    import win32com.client as win32
    
    c_pref, p_pref, c_items, p_items = extract_word_note_items(doc_path)
    
    if not output_path:
        base, ext = os.path.splitext(doc_path)
        output_path = f"{base}_附注编号已规范{ext}"
        
    shutil.copy2(doc_path, output_path)
    
    word = win32.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = False
    
    changes = []
    try:
        doc = word.Documents.Open(os.path.abspath(output_path))
        
        # 针对合并项重编号
        for seq, (_, old_num, subj, raw_text) in enumerate(c_items, start=1):
            target_text = f"{seq}、{subj}"
            if raw_text != target_text:
                rng = doc.Content
                find = rng.Find
                find.ClearFormatting()
                find.Replacement.ClearFormatting()
                if find.Execute(FindText=raw_text, ReplaceWith=target_text, Replace=1):
                    changes.append(f"[合并附注] 原【{raw_text}】 -> 规范重编为【{target_text}】")
                    
        # 针对母公司项重编号
        for seq, (_, old_num, subj, raw_text) in enumerate(p_items, start=1):
            target_text = f"{seq}、{subj}"
            if raw_text != target_text:
                rng = doc.Content
                find = rng.Find
                find.ClearFormatting()
                find.Replacement.ClearFormatting()
                if find.Execute(FindText=raw_text, ReplaceWith=target_text, Replace=1):
                    changes.append(f"[母公司附注] 原【{raw_text}】 -> 规范重编为【{target_text}】")
                    
        doc.Save()
        doc.Close(True)
    finally:
        word.Quit()
        
    return output_path, changes

def fix_financial_statements_note_numbers(excel_path, items_map, chapter_prefix='六、', output_excel_path=None):
    """
    一键补全或修复报表 Excel 中的附注编号
    """
    if not output_excel_path:
        base, ext = os.path.splitext(excel_path)
        output_excel_path = f"{base}_附注编号已修复{ext}"
        
    shutil.copy2(excel_path, output_excel_path)
    wb = openpyxl.load_workbook(output_excel_path)
    
    changes = []
    
    for sname in wb.sheetnames:
        if not any(k in sname for k in ['资产负债表', '利润表', '现金流量表']):
            continue
        ws = wb[sname]
        
        note_cols = []
        for r in range(1, 10):
            for c in range(1, ws.max_column + 1):
                val = str(ws.cell(r, c).value or '').strip()
                if val in ['附注', '附注编号', '附注号']:
                    note_cols.append(c)
        note_cols = sorted(list(set(note_cols)))
        if not note_cols:
            continue
            
        for r in range(5, ws.max_row + 1):
            for nc in note_cols:
                item_col = nc - 1
                item_val = str(ws.cell(r, item_col).value or '').strip()
                if not item_val:
                    continue
                clean_item = clean_subject_name(item_val)
                if not clean_item:
                    continue
                    
                matched_num = None
                for k, v in items_map.items():
                    if clean_item == k or (len(clean_item) >= 4 and clean_item in k) or (len(k) >= 4 and k in clean_item):
                        matched_num = v
                        break
                        
                if matched_num:
                    target_note_str = f"{chapter_prefix}{matched_num}" if chapter_prefix else str(matched_num)
                    current_note = str(ws.cell(r, nc).value or '').strip()
                    
                    if not current_note:
                        ws.cell(r, nc, target_note_str)
                        changes.append(f"【{sname}】科目【{clean_item}】（行 {r}）：空白 -> 自动补全为【{target_note_str}】")
                    elif current_note != target_note_str:
                        ws.cell(r, nc, target_note_str)
                        changes.append(f"【{sname}】科目【{clean_item}】（行 {r}）：原【{current_note}】 -> 修正为【{target_note_str}】")
                        
    wb.save(output_excel_path)
    return output_excel_path, changes
