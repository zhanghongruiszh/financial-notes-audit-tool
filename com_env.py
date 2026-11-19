# -*- coding: utf-8 -*-
"""
com_env.py — 冻结环境里调用 Office COM 前的环境消毒。

为什么需要
----------
PyInstaller onefile 运行时会把解包目录 `_MEIxxxx` 加到 PATH 和 DLL 搜索路径，
并留下 `_PYI_*` 之类的环境变量。这个进程再去 `DispatchEx('Word.Application')`
时，Word 作为子进程**继承**了这套环境，可能加载到 _MEI 里的同名 DLL
（vcruntime、msvcp、图形库等），结果就是 Word 起不来、或起来后"未能引发事件"。

症状：打包版里 Excel COM 正常，Word 一 Dispatch 就挂；源码版一切正常。

做法
----
调用 Office 之前把 `_MEI` 相关路径从 PATH 摘掉、把 PyInstaller 变量暂存，
调用完再恢复。源码运行时什么都不做。
"""

import os
import sys

_SAVED = {}


def sanitize():
    """摘掉 PyInstaller 留下的环境痕迹，返回是否做了改动。"""
    mei = getattr(sys, '_MEIPASS', None)
    if not mei:
        return False

    changed = False
    mei_norm = os.path.normcase(os.path.abspath(mei))

    path = os.environ.get('PATH', '')
    keep = []
    for part in path.split(os.pathsep):
        if not part:
            continue
        try:
            if os.path.normcase(os.path.abspath(part)) == mei_norm:
                changed = True
                continue
        except Exception:
            pass
        keep.append(part)
    if changed:
        _SAVED['PATH'] = path
        os.environ['PATH'] = os.pathsep.join(keep)

    for key in list(os.environ):
        if key.startswith('_PYI') or key in ('_MEIPASS2', '_MEIPASS'):
            _SAVED[key] = os.environ.pop(key)
            changed = True

    # 让 Windows 的 DLL 加载不再优先看当前工作目录
    if 'NoDefaultCurrentDirectoryInExePath' not in os.environ:
        _SAVED['NoDefaultCurrentDirectoryInExePath'] = None
        os.environ['NoDefaultCurrentDirectoryInExePath'] = '1'
        changed = True

    return changed


def restore():
    """恢复 sanitize() 改过的环境。"""
    for key, value in list(_SAVED.items()):
        if value is None:
            os.environ.pop(key, None)
        else:
            os.environ[key] = value
    _SAVED.clear()


class sanitized:
    """上下文管理器：with sanitized(): DispatchEx(...)"""

    def __enter__(self):
        self.changed = sanitize()
        return self

    def __exit__(self, *exc):
        restore()
        return False
