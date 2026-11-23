# -*- coding: utf-8 -*-
"""
selfcheck.py — 成果表与代码的自检。

检查项
------
1. 代码：重复的顶层函数/类定义（同一个名字定义两次，后者覆盖前者 —— 曾经踩过）
2. 代码：所有模块都能 import
3. 成果：必备工作表齐全
4. 成果：勾稽底稿里没有文本型「-」占用数字格
5. 成果：条件格式规则存在（有公式=绿、|差|>0.01=红）
6. 成果：报表勾稽列在最右（横向勾稽列不能压过去）
7. 成果：每条【勾稽差额】的公式都能解析、引用格都在本表范围、引用的都是数字格
8. 成果：汇总表「三、附注内部勾稽差异」条数与底稿红色一致

用法
----
    python tools/selfcheck.py <成果xlsx> [<成果xlsx> ...]
"""

import io
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from openpyxl.utils import column_index_from_string as C

HORIZ_HEADERS = ('横加', '余额-准备', '原值-折旧', '期初期末', '比例')
REF_RE = re.compile(r"'?([^'!]*)'?!?([A-Z]{1,2})(\d+)|(?<![A-Za-z0-9_])([A-Z]{1,2})(\d+)")


def check_code():
    out = []
    dup = []
    for fn in sorted(os.listdir('.')):
        if not fn.endswith('.py'):
            continue
        try:
            lines = io.open(fn, encoding='utf-8').read().split('\n')
        except Exception:
            continue
        seen = {}
        for i, ln in enumerate(lines, 1):
            m = re.match(r'^(def|class)\s+(\w+)', ln)
            if m:
                seen.setdefault(m.group(2), []).append(i)
        for k, v in seen.items():
            if len(v) > 1:
                dup.append((fn, k, v))
    out.append(('重复定义', dup, len(dup) == 0))
    return out


def check_imports():
    import importlib
    bad = []
    for fn in sorted(os.listdir('.')):
        if not fn.endswith('.py') or fn.startswith('_'):
            continue
        mod = fn[:-3]
        try:
            importlib.import_module(mod)
        except Exception as exc:
            bad.append((mod, str(exc)[:80]))
    return [('模块导入', bad, not bad)]


def _cells_with_formula(ws):
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith('='):
                yield cell


def check_workbook(path):
    out = []
    name = os.path.basename(path)

    # 3. 工作表
    wb = openpyxl.load_workbook(path)
    need = ['汇总表', '勾稽底稿', '完整性', '原始数据']
    missing = [s for s in need if s not in wb.sheetnames]
    out.append((f'{name} 工作表', missing, not missing))

    ws = wb['勾稽底稿']

    # 4. 文本「-」
    dashes = []
    for row in ws.iter_rows(min_col=2):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip() in ('-', '—', '——', '--'):
                dashes.append(cell.coordinate)
    out.append((f'{name} 文本破折号', dashes[:5], not dashes))

    # 5. 条件格式
    cf = list(ws.conditional_formatting)
    nrules = sum(len(r.rules) for r in cf)
    out.append((f'{name} 条件格式', nrules, nrules >= 2))

    # 6. 报表列最右
    bad_order = 0
    for r in range(1, ws.max_row + 1):
        horiz, rep = [], []
        for c in range(1, 30):
            v = ws.cell(r, c).value
            if not isinstance(v, str):
                continue
            if v in HORIZ_HEADERS:
                horiz.append(c)
            if v.startswith('【报表勾稽'):
                rep.append(c)
        if horiz and rep and max(horiz) > min(rep):
            bad_order += 1
    out.append((f'{name} 报表列在最右', bad_order, bad_order == 0))

    # 7. 勾稽差额公式
    problems = []
    nonzeros = []
    nchk = 0
    for row in range(1, ws.max_row + 1):
        if ws.cell(row, 1).value != '【勾稽差额】':
            continue
        nchk += 1
        for c in range(2, 30):
            cell = ws.cell(row, c)
            v = cell.value
            if not isinstance(v, str) or not v.startswith('='):
                continue
            if 'None' in v:
                problems.append((cell.coordinate, 'NONE', v[:40]))
                continue
            for m in REF_RE.finditer(v):
                col = m.group(2) or m.group(4)
                rr = m.group(3) or m.group(5)
                if not col or not rr:
                    continue
                rr = int(rr)
                if rr < 1 or rr > ws.max_row:
                    problems.append((cell.coordinate, 'OUT', f'{col}{rr}'))

        # 重算公式，统计"真的有差异"的格数
        for c in range(2, 30):
            cell = ws.cell(row, c)
            v = cell.value
            if not isinstance(v, str) or not v.startswith('='):
                continue
            val = _eval_formula(v, ws)
            if val is None:
                problems.append((cell.coordinate, 'EVAL', v[:40]))
            elif abs(val) > 0.01:
                nonzeros.append((cell.coordinate, val))
    out.append((f'{name} 勾稽差额行', f'{nchk} 行，异常 {len(problems)}',
                not problems))
    # 横向勾稽（表右侧的「期初期末」「横加」等列）也要重算
    horiz_nonzero = []
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if not isinstance(v, str) or not v.startswith('='):
                continue
            if ws.cell(cell.row, 1).value == '【勾稽差额】':
                continue
            val = _eval_formula(v, ws)
            if val is not None and abs(val) > 0.01:
                horiz_nonzero.append((cell.coordinate, val))
    total_nz = len(nonzeros) + len(horiz_nonzero)
    out.append((f'{name} 公式重算非零格',
                f'纵向 {len(nonzeros)} + 横向 {len(horiz_nonzero)} = {total_nz}', True))
    out.append((f'{name} 非零明细', (nonzeros + horiz_nonzero)[:8], True))
    if problems:
        out.append((f'{name} 异常示例', problems[:5], False))

    # 8. 汇总表差异条数
    sm = wb['汇总表']
    declared = None
    for r in range(1, sm.max_row + 1):
        v = sm.cell(r, 1).value
        if isinstance(v, str) and v.startswith('三、附注内部勾稽差异'):
            m = re.search(r'（(\d+)\s*条）', v)
            declared = int(m.group(1)) if m else 0
    reds = 0
    if declared:
        for r in range(1, sm.max_row + 1):
            v = sm.cell(r, 5).value
            if isinstance(v, (int, float)):
                reds += 1
    out.append((f'{name} 汇总差异', f'声明 {declared}', declared is not None))
    # 最强的一条：Excel 公式重算出来的差异数，必须等于汇总表声明的差异数
    if declared is not None:
        out.append((f'{name} 公式与汇总一致',
                    f'公式 {total_nz} 条 / 汇总 {declared} 条',
                    total_nz == declared))
    return out



def ref_to_val(expr, ws):
    out = []
    i = 0
    for m in re.finditer(r'([A-Z]{1,2})(\d+)', expr):
        piece = expr[i:m.start()]
        if not re.fullmatch(r'[\d\s+\-*/().]*', piece):
            return None
        out.append(piece)
        col = m.group(1)
        rr = m.group(2)
        if not col or not rr:
            return None
        v = ws.cell(int(rr), C(col)).value
        if v is None:
            out.append('0.0')
        elif isinstance(v, (int, float)):
            out.append(repr(float(v)))
        else:
            return None
        i = m.end()
    tail = expr[i:]
    if not re.fullmatch(r'[\d\s+\-*/().]*', tail):
        return None
    out.append(tail)
    return ''.join(out)


def _eval_formula(text, ws):
    expr = text[1:]
    if 'None' in expr:
        return None
    src = ref_to_val(expr, ws)
    if src is None:
        return None
    try:
        return float(eval(src, {'__builtins__': {}}, {}))
    except Exception:
        return None


def main():
    targets = sys.argv[1:]
    if not targets:
        print('用法：python tools/selfcheck.py <成果xlsx> ...')
        return 2
    results = check_code() + check_imports()
    for p in targets:
        if os.path.exists(p):
            results += check_workbook(p)
        else:
            results.append((p, '文件不存在', False))

    ok = 0
    print('=' * 68)
    for item, detail, passed in results:
        mark = 'PASS' if passed else 'FAIL'
        print(f'[{mark}] {item}：{detail}')
        ok += 1 if passed else 0
    print('=' * 68)
    print(f'共 {len(results)} 项，通过 {ok}，失败 {len(results) - ok}')
    return 0 if ok == len(results) else 1


if __name__ == '__main__':
    raise SystemExit(main())
