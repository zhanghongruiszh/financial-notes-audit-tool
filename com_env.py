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

import hashlib
import os
import shutil
import sys
import tempfile

_SAVED = {}

STAGING_DIR = 'notes_conv'


def _is_ascii(path):
    try:
        path.encode('ascii')
        return True
    except UnicodeEncodeError:
        return False


def _candidates(preferred_dir=None):
    """按优先级给出可用的暂存目录。

    优先级：调用方给的工作目录 → 当前工作目录 → 程序目录 → 系统临时目录。
    只接受**纯 ASCII 且确实可写**的路径——中文路径会让 Office 打不开，
    而受限环境里 AppData\\Temp 也可能不可写。
    """
    out = []
    if preferred_dir:
        out.append(os.path.join(os.path.abspath(preferred_dir), '_conv'))
    try:
        out.append(os.path.join(os.getcwd(), '_conv'))
    except Exception:
        pass
    try:
        out.append(os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), '_conv'))
    except Exception:
        pass
    out.append(os.path.join(tempfile.gettempdir(), STAGING_DIR))
    local = os.environ.get('LOCALAPPDATA')
    if local:
        out.append(os.path.join(local, 'Temp', STAGING_DIR))
    return out


def ascii_staging(path, suffix=None, preferred_dir=None):
    """把文件摆到一个纯 ASCII 的暂存路径，供 Office 打开。

    打包版里把中文路径直接交给 Word/Excel，会出现：
        Word：Word 未能引发事件 / 卡住不返回
        Excel：不能访问 INetCache\\Content.MSO\\
    同一文件复制成纯英文路径就完全正常——是路径编码在冻结环境里被破坏。

    返回 (可交给 Office 的路径, 是否是复制出来的副本)。
    """
    path = os.path.abspath(path)
    if _is_ascii(path):
        return path, False

    suffix = suffix or os.path.splitext(path)[1]
    digest = hashlib.md5(path.encode('utf-8')).hexdigest()[:10]

    for base in _candidates(preferred_dir):
        if not _is_ascii(base):
            continue
        try:
            os.makedirs(base, exist_ok=True)
            dst = os.path.join(base, f'src_{digest}{suffix}')
            if (not os.path.exists(dst)
                    or os.path.getmtime(dst) < os.path.getmtime(path)):
                shutil.copy2(path, dst)
            return dst, True
        except Exception:
            continue

    return path, False


def staging_dir(preferred_dir=None):
    """与 ascii_staging 同目录，用于放转换结果。"""
    for base in _candidates(preferred_dir):
        if not _is_ascii(base):
            continue
        try:
            os.makedirs(base, exist_ok=True)
            probe = os.path.join(base, '.w')
            with open(probe, 'w') as f:
                f.write('x')
            os.remove(probe)
            return base
        except Exception:
            continue
    return None


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
