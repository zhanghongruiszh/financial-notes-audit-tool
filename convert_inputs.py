# -*- coding: utf-8 -*-
"""
convert_inputs.py — 把导入的 .doc / .xls 统一转成 .docx / .xlsx。

为什么单独做一步
----------------
`.docx` / `.xlsx` 是**直接读取**的，不经过 Office；只有 `.doc` / `.xls` 才要
先调 Word / Excel 转换。而 Office 自动化在远程桌面、服务、受限环境里经常不可用
（表现为 Word 卡住、Excel 报 INetCache 无权限）。

所以提供"先转换、再勾稽"的用法：
    · 转换一次，之后都读 docx/xlsx，快且稳
    · 转换结果就放在原文件旁边，用户可以自己看、自己拿走

对外接口
--------
    convert_one(path)        -> (新路径, 说明)
    convert_many(paths)      -> [(原路径, 新路径或 None, 说明)]
"""

import os

import com_env
import docx_ir

DOC_TARGET = {'.doc': '.docx'}
XLS_TARGET = {'.xls': '.xlsx'}


def target_path(path):
    """给出转换后应放的位置：与原文件同目录、同名、换扩展名。"""
    base, ext = os.path.splitext(os.path.abspath(path))
    low = ext.lower()
    if low in DOC_TARGET:
        return base + DOC_TARGET[low]
    if low in XLS_TARGET:
        return base + XLS_TARGET[low]
    return None


def needs_convert(path):
    return os.path.splitext(str(path))[1].lower() in ('.doc', '.xls')


def convert_one(path, log=print):
    """把一个 .doc/.xls 转成同目录的 .docx/.xlsx。返回 (新路径, 说明)。"""
    path = os.path.abspath(path)
    if not needs_convert(path):
        return path, '无需转换'

    dst = target_path(path)
    if os.path.exists(dst) and os.path.getmtime(dst) >= os.path.getmtime(path):
        log(f'      已有转换结果，直接用：{os.path.basename(dst)}')
        return dst, '复用了已有转换结果'

    log(f'      转换中：{os.path.basename(path)} → {os.path.basename(dst)}')
    work = os.path.join(os.path.dirname(dst), '_conv')
    if os.path.splitext(path)[1].lower() == '.doc':
        made = docx_ir.ensure_docx(path, work_dir=work)
    else:
        import tb_reader
        made = tb_reader.ensure_xlsx(path, work_dir=work)

    os.makedirs(os.path.dirname(dst) or '.', exist_ok=True)
    if os.path.abspath(made) != os.path.abspath(dst):
        import shutil
        shutil.copy2(made, dst)
    return dst, '已转换'


def convert_many(paths, log=print):
    """批量转换，返回 [(原路径, 新路径或 None, 说明)]。"""
    out = []
    for p in paths:
        if not p:
            continue
        try:
            new, msg = convert_one(p, log=log)
            out.append((p, new, msg))
        except Exception as exc:            # noqa: BLE001
            log(f'      转换失败：{os.path.basename(p)}：{exc}')
            out.append((p, None, str(exc)))
    return out
