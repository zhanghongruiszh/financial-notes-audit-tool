# -*- coding: utf-8 -*-
"""
财务报表附注自动化勾稽系统 — 桌面入口

流程
----
在界面上选三个文件（附注 Word、报表 Excel、TB 试算表）→ 一键勾稽 →
生成 `<项目名>check_v<版本号>/` 文件夹，里面放：
    附注 Word、报表 Excel、TB 试算表（备查）
    <项目名>check_v<版本号>.xlsx（勾稽成果）

另附两个独立功能：
    · Word 附注编号体检与一键规范（改成自动编号，从 1 连续）
    · 打开最近一次的输出文件夹
"""

import os
import sys
import tempfile
import threading
import tkinter as tk
import traceback
import webbrowser
from tkinter import filedialog, messagebox, ttk

import numbering_fixer

GITHUB_URL = 'https://github.com/zhanghongruiszh/financial-notes-audit-tool'
AUTHOR_EMAIL = 'henryzhang825@gmail.com'

# 打包成 exe 后，工作目录可能是只读的（比如放在 Program Files），
# 中间转换文件统一放系统临时目录。
WORK_DIR = os.path.join(tempfile.gettempdir(), 'notes_audit_work')
os.makedirs(WORK_DIR, exist_ok=True)


def _co_init():
    """工作线程里调 Office 前初始化 COM（STA）。

    Word/Excel 自动化要求调用线程是 STA。界面里这些活都在工作线程跑，
    不显式初始化的话，打包环境里很容易卡死在 Documents.Open。
    """
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass


class AuditCheckApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('财务报表附注自动化勾稽系统 v3.0')
        self.geometry('980x720')
        self.minsize(900, 640)
        self.last_out_dir = None
        self.center_window()

        self.style = ttk.Style()
        try:
            self.style.theme_use('vista')
        except tk.TclError:
            pass

        # 首次运行把内置模板骨架放到程序目录，之后以磁盘上的为准（可自行替换）
        try:
            import template_lib
            template_lib.seed_if_empty()
        except Exception:
            pass

        self.create_widgets()

    def center_window(self):
        self.update_idletasks()
        w, h = 980, 720
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2) - 20
        self.geometry(f'{w}x{h}+{x}+{y}')

    # ------------------------------------------------------------ 界面

    def create_widgets(self):
        main = ttk.Frame(self, padding='15 12 15 10')
        main.pack(fill=tk.BOTH, expand=True)

        tk.Label(main, text='财务报表附注自动化勾稽系统',
                 font=('微软雅黑', 14, 'bold'), fg='#1F497D').pack(anchor=tk.W)

        grp_files = ttk.LabelFrame(main, text=' 提交原始文件 ', padding='12 10 12 10')
        grp_files.pack(fill=tk.X, pady=(8, 10))

        self.var_notes = tk.StringVar()
        self.var_report = tk.StringVar()
        self.var_tb = tk.StringVar()

        rows = [
            ('1. 报告附注:', self.var_notes, self.browse_notes,
             [('Word 附注', '*.doc *.docx'), ('所有文件', '*.*')]),
            ('2. 财务报表:', self.var_report, self.browse_report,
             [('Excel 报表', '*.xlsx *.xls'), ('所有文件', '*.*')]),
            ('3. TB试算表:', self.var_tb, self.browse_tb,
             [('Excel 底稿', '*.xls *.xlsx'), ('所有文件', '*.*')]),
        ]
        for i, (label, var, cmd, _) in enumerate(rows):
            ttk.Label(grp_files, text=label).grid(row=i, column=0, sticky=tk.W, pady=5)
            ttk.Entry(grp_files, textvariable=var, width=70).grid(
                row=i, column=1, sticky=tk.EW, padx=8, pady=5)
            ttk.Button(grp_files, text='选择...', command=cmd).grid(
                row=i, column=2, sticky=tk.W, pady=5)
        grp_files.columnconfigure(1, weight=1)

        grp_cfg = ttk.LabelFrame(main, text=' 归档设置 ', padding='12 8 12 8')
        grp_cfg.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(grp_cfg, text='项目名:').grid(row=0, column=0, sticky=tk.W, pady=5)
        self.var_project = tk.StringVar()
        ttk.Entry(grp_cfg, textvariable=self.var_project, width=30).grid(
            row=0, column=1, sticky=tk.W, padx=8, pady=5)
        ttk.Label(grp_cfg, text='（留空则从文件名自动识别）',
                  foreground='#666666', font=('微软雅黑', 8)).grid(
            row=0, column=2, sticky=tk.W, pady=5)

        ttk.Label(grp_cfg, text='归档目录:').grid(row=1, column=0, sticky=tk.W, pady=5)
        self.var_archive = tk.StringVar(value=os.path.abspath('.'))
        ttk.Entry(grp_cfg, textvariable=self.var_archive, width=60).grid(
            row=1, column=1, columnspan=2, sticky=tk.EW, padx=8, pady=5)
        ttk.Button(grp_cfg, text='选择目录...', command=self.browse_archive).grid(
            row=1, column=3, sticky=tk.W, pady=5)
        grp_cfg.columnconfigure(1, weight=1)

        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill=tk.X, pady=(0, 10))

        self.btn_run = tk.Button(
            btn_frame, text='开始自动化全面勾稽', font=('微软雅黑', 10, 'bold'),
            bg='#1F497D', fg='white', activebackground='#153358',
            activeforeground='white', relief=tk.RAISED, cursor='hand2',
            padx=14, pady=7, command=self.on_run)
        self.btn_run.pack(side=tk.LEFT, padx=(0, 10))

        tk.Button(btn_frame, text='Word附注编号体检与修复', font=('微软雅黑', 9),
                  bg='#EBF1F5', fg='#1F497D', cursor='hand2', padx=10, pady=6,
                  command=self.on_fix_numbering).pack(side=tk.LEFT, padx=(0, 10))

        tk.Button(btn_frame, text='报表附注号一键修复与补全', font=('微软雅黑', 9),
                  bg='#EBF1F5', fg='#1F497D', cursor='hand2', padx=10, pady=6,
                  command=self.on_fix_report_numbers).pack(side=tk.LEFT, padx=(0, 10))

        tk.Button(btn_frame, text='模板库', font=('微软雅黑', 9),
                  bg='#EBF1F5', fg='#1F497D', cursor='hand2', padx=10, pady=6,
                  command=self.on_manage_templates).pack(side=tk.LEFT, padx=(0, 10))

        tk.Button(btn_frame, text='打开输出文件夹', font=('微软雅黑', 9),
                  bg='#F2F2F2', fg='#333333', cursor='hand2', padx=10, pady=6,
                  command=self.open_out_dir).pack(side=tk.RIGHT)

        # 页脚：先于日志区 pack 且贴底部，保证窗口再小也不会被挤出视野
        footer = ttk.Frame(main)
        footer.pack(side=tk.BOTTOM, fill=tk.X, pady=(6, 0))
        tk.Label(footer, text='开源项目地址：', font=('微软雅黑', 8)).pack(side=tk.LEFT)
        link = tk.Label(footer, text=GITHUB_URL, font=('微软雅黑', 8, 'underline'),
                        fg='#0969DA', cursor='hand2')
        link.pack(side=tk.LEFT)
        link.bind('<Button-1>', lambda e: webbrowser.open_new_tab(GITHUB_URL))
        tk.Label(footer, text=f'问题反馈邮箱：{AUTHOR_EMAIL}',
                 font=('微软雅黑', 8), fg='#1F497D').pack(side=tk.RIGHT)

        grp_log = ttk.LabelFrame(main, text=' 执行日志 ', padding='10 8 10 8')
        grp_log.pack(fill=tk.BOTH, expand=True)

        self.txt_log = tk.Text(grp_log, wrap=tk.WORD, font=('Consolas', 9),
                               bg='#F8F9FA', fg='#262626')
        self.txt_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb = ttk.Scrollbar(grp_log, orient=tk.VERTICAL, command=self.txt_log.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt_log.config(yscrollcommand=sb.set)

        self.log('[系统状态] 就绪。请选择报告附注、财务报表与 TB 试算表。')
        self.log('[说明] 成果按 <项目名>check_v<版本号> 建文件夹，'
                 '每版都放入附注、报表、TB 与 check 结果四个文件。')

    # ------------------------------------------------------------ 工具

    def log(self, msg):
        self.txt_log.insert(tk.END, str(msg) + '\n')
        self.txt_log.see(tk.END)
        self.update_idletasks()

    def _pick(self, var, title, filetypes):
        path = filedialog.askopenfilename(title=title, filetypes=filetypes)
        if path:
            var.set(path)

    def browse_notes(self):
        self._pick(self.var_notes, '选择报告附注文件',
                   [('Word 附注', '*.doc *.docx'), ('所有文件', '*.*')])

    def browse_report(self):
        self._pick(self.var_report, '选择财务报表文件',
                   [('Excel 报表', '*.xlsx *.xls'), ('所有文件', '*.*')])

    def browse_tb(self):
        self._pick(self.var_tb, '选择 TB 试算平衡表',
                   [('Excel 底稿', '*.xls *.xlsx'), ('所有文件', '*.*')])

    def browse_archive(self):
        path = filedialog.askdirectory(title='选择归档目录')
        if path:
            self.var_archive.set(path)

    def open_out_dir(self):
        folder = self.last_out_dir or self.var_archive.get()
        if folder and os.path.isdir(folder):
            os.startfile(folder)
        else:
            messagebox.showinfo('提示', '还没有生成过成果文件夹，请先执行勾稽。')

    def out_target_dir(self, source_path):
        """编号治理产物的落点。

        已经跑过勾稽 → 放进那次的项目输出文件夹（这样一个文件夹里就有 4~6 个文件：
        附注、报表、TB、check 成果，加上编号规范后的附注、编号修复后的报表）；
        还没跑过勾稽 → 先放在源文件旁边，等勾稽时会一并归档进输出文件夹。
        """
        if self.last_out_dir and os.path.isdir(self.last_out_dir):
            return self.last_out_dir
        return os.path.dirname(os.path.abspath(source_path))

    # ------------------------------------------------------------ 勾稽

    def on_run(self):
        notes = self.var_notes.get().strip()
        report = self.var_report.get().strip()
        tb = self.var_tb.get().strip()
        root = self.var_archive.get().strip() or os.path.abspath('.')
        project = self.var_project.get().strip()

        if not notes or not os.path.exists(notes):
            messagebox.showwarning('提示', '请先选择有效的报告附注文件。')
            return
        if not report or not os.path.exists(report):
            messagebox.showwarning('提示', '请先选择有效的财务报表文件。')
            return
        if not tb or not os.path.exists(tb):
            messagebox.showwarning('提示', '请先选择有效的 TB 试算表。')
            return

        self.btn_run.config(state=tk.DISABLED, text='勾稽运算中...')
        self.log('\n' + '=' * 60)

        def work():
            _co_init()
            try:
                import pipeline
                res = pipeline.run(notes, report, tb, archive_root=root,
                                   project_name=project, log=self.log)
                self.last_out_dir = res['out_dir']
                stats = res['stats']
                self.log('[完成] 成果已生成。')
                messagebox.showinfo(
                    '勾稽完成',
                    f"项目【{res['project']}】第 {res['version']} 版已生成。\n\n"
                    f"附注内部勾稽 {res['checks']} 条："
                    f"通过 {stats.get('pass', 0)}，有差异 {stats.get('fail', 0)}\n"
                    f"报表勾稽 {res['cross']} 条\n\n"
                    f"文件：\n{res['out_xlsx']}")
            except Exception as exc:
                self.log('[错误] ' + str(exc))
                self.log(traceback.format_exc())
                messagebox.showerror('勾稽失败', f'{exc}')
            finally:
                self.btn_run.config(state=tk.NORMAL, text='开始自动化全面勾稽')

        threading.Thread(target=work, daemon=True).start()

    # ------------------------------------------------------------ 模板库

    def on_manage_templates(self):
        """模板管理：查看 / 导入 / 删除 / 从 Word 重建骨架。"""
        import template_lib

        win = tk.Toplevel(self)
        win.title('模板库管理')
        win.geometry('760x430')
        win.transient(self)

        top = ttk.Frame(win, padding='10 8 10 4')
        top.pack(fill=tk.X)
        ttk.Label(top, text='模板只用于"加速识别"：有相近模板时补全列角色判断，'
                            '没有模板也能正常勾稽。',
                  foreground='#666666').pack(anchor=tk.W)
        ttk.Label(top, text=f'模板库目录：{template_lib.library_dir()}',
                  foreground='#1F497D').pack(anchor=tk.W, pady=(4, 0))

        frame = ttk.Frame(win, padding='10 4 10 4')
        frame.pack(fill=tk.BOTH, expand=True)
        cols = ('name', 'subs', 'tables', 'source')
        tree = ttk.Treeview(frame, columns=cols, show='headings', height=12)
        for key, text, width in (('name', '模板名称', 230), ('subs', '子项目', 70),
                                 ('tables', '表格数', 70), ('source', '来源文件', 300)):
            tree.heading(key, text=text)
            tree.column(key, width=width, anchor=tk.W)
        tree.pack(fill=tk.BOTH, expand=True)

        def refresh():
            tree.delete(*tree.get_children())
            for sk in template_lib.load_library():
                subs = sum(len(c.get('sub_items', [])) for c in sk.get('chapters', []))
                tabs = sum(len(i.get('tables', []))
                           for c in sk.get('chapters', []) for i in c.get('sub_items', []))
                tree.insert('', 'end', values=(sk.get('name', ''),
                                               subs, tabs, sk.get('source', '')),
                            tags=(sk.get('_file', ''),))

        def selected_file():
            sel = tree.selection()
            if not sel:
                messagebox.showinfo('提示', '请先选中一个模板。', parent=win)
                return None
            return tree.item(sel[0], 'tags')[0]

        def do_import():
            path = filedialog.askopenfilename(
                title='选择要导入的模板（Word 附注模板）',
                filetypes=[('Word 模板', '*.doc *.docx'), ('所有文件', '*.*')])
            if not path:
                return

            def work():
                _co_init()
                try:
                    self.log(f'\n[模板] 正在提取骨架：{os.path.basename(path)} ...')
                    sk = template_lib.build_skeleton(path, work_dir=WORK_DIR)
                    saved = template_lib.save_skeleton(sk)
                    self.log(f'[模板] 已加入：{os.path.basename(saved)}')
                    win.after(0, refresh)
                    win.after(0, lambda: messagebox.showinfo(
                        '完成', f'模板已导入：\n{os.path.basename(saved)}', parent=win))
                except Exception as exc:
                    self.log('[模板] 导入失败：' + str(exc))
                    win.after(0, lambda: messagebox.showerror('导入失败', str(exc), parent=win))

            threading.Thread(target=work, daemon=True).start()

        def do_rebuild():
            def work():
                _co_init()
                try:
                    self.log('\n[模板] 从模板库目录里的 Word 文件重建骨架 ...')
                    done = template_lib.rebuild_from_docs(work_dir=WORK_DIR, log=self.log)
                    win.after(0, refresh)
                    msg = (f'已重建 {len(done)} 个模板骨架。' if done
                           else '模板库目录里没有 .doc/.docx 文件。')
                    win.after(0, lambda: messagebox.showinfo('重建完成', msg, parent=win))
                except Exception as exc:
                    self.log('[模板] 重建失败：' + str(exc))
                    win.after(0, lambda: messagebox.showerror('重建失败', str(exc), parent=win))

            threading.Thread(target=work, daemon=True).start()

        def do_delete():
            name = selected_file()
            if not name:
                return
            if not messagebox.askyesno('确认', f'删除模板「{name}」？', parent=win):
                return
            try:
                os.remove(os.path.join(template_lib.library_dir(), name))
                self.log(f'[模板] 已删除 {name}')
                refresh()
            except Exception as exc:
                messagebox.showerror('删除失败', str(exc), parent=win)

        def do_open():
            os.startfile(template_lib.library_dir())

        bottom = ttk.Frame(win, padding='10 4 10 10')
        bottom.pack(fill=tk.X)
        ttk.Button(bottom, text='导入模板...', command=do_import).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(bottom, text='从 Word 重建骨架',
                   command=do_rebuild).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(bottom, text='删除选中', command=do_delete).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(bottom, text='打开模板库文件夹', command=do_open).pack(side=tk.LEFT)
        ttk.Button(bottom, text='关闭', command=win.destroy).pack(side=tk.RIGHT)

        refresh()

    # ------------------------------------------------------------ 报表附注号

    def on_fix_report_numbers(self):
        """按附注的实际编号，补全/修正报表主表的「附注」列。"""
        doc = self.var_notes.get().strip()
        report = self.var_report.get().strip()
        if not doc or not os.path.exists(doc):
            messagebox.showwarning('提示', '请先在【报告附注】处选择 Word 文件。')
            return
        if not report or not os.path.exists(report):
            messagebox.showwarning('提示', '请先在【财务报表】处选择 Excel 文件。')
            return

        def work():
            _co_init()
            try:
                import notes_parser
                import report_numbering
                self.log(f'\n[报表编号] 读取附注：{os.path.basename(doc)}')
                model = notes_parser.parse_notes(doc, work_dir=WORK_DIR)
                for ch in model.chapters:
                    self.log(f'  章<{ch.name}> 编号「{ch.number_text or "无"}」，'
                             f'子项目 {len(ch.sub_items)} 个')

                base, ext = os.path.splitext(report)
                out = os.path.join(self.out_target_dir(report),
                                   f'{os.path.basename(base)}_附注编号已修复{ext}')
                out, changes = report_numbering.fix_report_numbers(
                    report, model, out, log=self.log)
                for c in changes[:30]:
                    self.log('   ' + c)
                if len(changes) > 30:
                    self.log(f'   ... 另 {len(changes) - 30} 处')
                if not changes:
                    self.log('[报表编号] 没有需要补全或修正的附注号。')
                messagebox.showinfo(
                    '完成',
                    f'共处理 {len(changes)} 处附注号。\n\n文件：\n{out}')
            except Exception as exc:
                self.log('[错误] ' + str(exc))
                self.log(traceback.format_exc())
                messagebox.showerror('报表附注号处理失败', str(exc))

        threading.Thread(target=work, daemon=True).start()

    # ------------------------------------------------------------ 编号治理

    def on_fix_numbering(self):
        doc = self.var_notes.get().strip()
        if not doc or not os.path.exists(doc):
            messagebox.showwarning('提示', '请先在【报告附注】处选择 Word 文件。')
            return

        def work():
            _co_init()
            try:
                self.log(f'\n[编号] 体检：{os.path.basename(doc)}')
                import notes_parser
                model = notes_parser.parse_notes(doc, work_dir=WORK_DIR)
                issues = numbering_fixer.diagnose(model)
                bad = [i for i in issues if i.problem != '正常']
                self.log(f"[编号] 共 {len(issues)} 个子项目，其中异常 {len(bad)} 处")
                for i in bad[:20]:
                    self.log(f'   {i.line()}')
                if not bad:
                    self.log('[编号] 编号已连续，无需修复。')
                    messagebox.showinfo('编号体检', '未发现编号异常，无需修复。')
                    return

                if not messagebox.askyesno(
                        '编号体检',
                        f'发现 {len(bad)} 处编号异常。\n\n是否生成修好的副本？\n'
                        '（会新建文件，绝不覆盖原件）'):
                    return

                out = os.path.join(
                    self.out_target_dir(doc),
                    os.path.basename(os.path.splitext(doc)[0]) + '_编号已规范.docx')
                _, _, detail = numbering_fixer.renumber_document(model, out, log=self.log)
                ok = all(v[0] for v in detail.values())
                self.log(f'[编号] 复核：{"全部连续 ✓" if ok else "仍有不连续 ✗"}')
                messagebox.showinfo(
                    '完成',
                    ('编号已规范，复核连续 ✓' if ok else '编号已处理，但复核发现不连续，请检查') +
                    f'\n\n文件：\n{out}')
            except Exception as exc:
                self.log('[错误] ' + str(exc))
                self.log(traceback.format_exc())
                messagebox.showerror('编号处理失败', str(exc))

        threading.Thread(target=work, daemon=True).start()


def selftest(out_path):
    """打包后的自检：在无交互环境下验证依赖与界面能否构建。

    exe 是窗口程序，直接启动会一直开着、无法在自动化里判断成功与否，
    所以提供一个 `--selftest <结果文件>` 开关：构建一遍界面 → 销毁 → 写结果 → 退出。
    """
    lines = []
    ok = True
    try:
        app = AuditCheckApp()
        app.update_idletasks()
        lines.append(f'界面构建成功，窗口 {app.winfo_width()}x{app.winfo_height()}')
        # 页脚（开源地址/邮箱）是否真的排进了窗口
        top = app.winfo_rooty() + app.winfo_height()
        footer_ok = False

        def walk(w):
            nonlocal footer_ok
            for c in w.winfo_children():
                try:
                    text = str(c.cget('text'))
                except Exception:
                    text = ''
                if 'github.com' in text:
                    f = c.master
                    footer_ok = f.winfo_rooty() + f.winfo_height() <= top
                walk(c)

        walk(app)
        lines.append(f'页脚可见：{footer_ok}')
        app.destroy()

        import openpyxl                      # noqa: F401
        import docx                          # noqa: F401
        import win32com.client               # noqa: F401
        lines.append('第三方依赖导入 OK（openpyxl / python-docx / pywin32）')

        import cross_checks                  # noqa: F401
        import docx_ir                       # noqa: F401
        import disclosure_checks             # noqa: F401
        import docx_numbering                # noqa: F401
        import note_cross_checks             # noqa: F401
        import note_number_checks            # noqa: F401
        import notes_parser                  # noqa: F401
        import numbering_fixer               # noqa: F401
        import pipeline                      # noqa: F401
        import report_numbering              # noqa: F401
        import report_reader                 # noqa: F401
        import rules                         # noqa: F401
        import table_semantics               # noqa: F401
        import table_types                   # noqa: F401
        import tb_checks                     # noqa: F401
        import tb_reader                     # noqa: F401
        import template_lib                  # noqa: F401
        import word_to_excel_extractor       # noqa: F401
        import writer                        # noqa: F401
        lines.append('业务模块导入 OK')
        lines.append('OK')
    except Exception:
        ok = False
        lines.append('FAIL')
        lines.append(traceback.format_exc())

    try:
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines))
    except Exception:
        pass
    return 0 if ok else 1


def headless_run(notes, report, tb, out_dir, log_path):
    """无界面跑一遍完整流程，把日志写到文件。

    给打包后的 exe 做端到端验证用（GUI 里的按钮走的是同一套 pipeline）。
    """
    lines = []

    def log(msg):
        lines.append(str(msg))

    try:
        # 冻结版里主线程默认不是 STA，Word COM 会卡在 DispatchEx；
        # 显式初始化一次（GUI 的按钮线程里也做了同样的事）。
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pass

    try:
        import pipeline
        log('[headless] 开始')
        result = pipeline.run(notes, report, tb, archive_root=out_dir, log=log)
        log('[headless] 完成：' + str(result.get('out_xlsx')))
        ok = True
    except Exception:
        import traceback
        log('[headless] 失败')
        log(traceback.format_exc())
        ok = False

    try:
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines))
    except Exception:
        pass
    return 0 if ok else 1


def com_probe(src_doc, log_path):
    """逐步探测 Word COM 在打包环境里卡在哪一步（诊断用）。"""
    lines = []

    def step(msg):
        lines.append(msg)
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines))

    import os as _os
    src = _os.path.abspath(src_doc)
    step('1 开始')

    def try_variant(name, fn):
        try:
            fn()
            step(f'   [{name}] OK')
            return True
        except Exception as exc:
            step(f'   [{name}] 失败: {exc}')
            return False

    # A: 显式 STA + 可见 Word
    def a():
        import pythoncom
        import win32com.client as win32
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        word = win32.DispatchEx('Word.Application')
        word.Visible = True
        doc = word.Documents.Open(src)
        doc.Close(False)
        word.Quit()

    # B: 不可见 + 只传文件名
    def b():
        import win32com.client as win32
        word = win32.DispatchEx('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(FileName=src)
        doc.Close(False)
        word.Quit()

    # C: 先打开文档再设属性
    def c():
        import win32com.client as win32
        word = win32.Dispatch('Word.Application')
        doc = word.Documents.Open(src, ReadOnly=True)
        doc.Close(False)
        word.Quit()

    step('2 基本 COM / Excel')
    try:
        import win32com.client as win32
        fso = win32.Dispatch('Scripting.FileSystemObject')
        step('   FSO OK: C盘存在=%s' % fso.DriveExists('C'))
    except Exception as exc:
        step('   FSO 失败: %s' % (exc,))
    try:
        excel = win32.DispatchEx('Excel.Application')
        step('   Excel DispatchEx OK')
        excel.Quit()
        step('   Excel Quit OK')
    except Exception as exc:
        step('   Excel 失败: %s' % (exc,))

    step('3 Word 分步（先消毒环境）')
    try:
        import com_env
        step('   环境消毒: %s' % com_env.sanitize())
    except Exception as exc:
        step('   环境消毒失败: %s' % (exc,))

    def one_word(fn, name):
        """每个变体都开一个干净的 Word，避免上一次失败污染下一次。"""
        word = None
        try:
            word = win32.DispatchEx('Word.Application')
            fn(word)
            step(f'   [{name}] OK')
            return True
        except Exception as exc:
            step(f'   [{name}] 失败: {exc}')
            return False
        finally:
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass

    variants = [
        ('A 可见后打开', lambda w: (setattr(w, 'Visible', True), w.Documents.Open(src))[1]),
        ('B 只读+不带确认', lambda w: w.Documents.Open(
            src, ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False)),
        ('C Open 带 Visible=False', lambda w: w.Documents.Open(src, Visible=False)),
        ('D 关掉告警再开', lambda w: (setattr(w, 'DisplayAlerts', 0),
                                     w.Documents.Open(src))[1]),
    ]
    results = [(name, one_word(fn, name)) for name, fn in variants]
    step('4 结果: ' + str(results))
    return 0 if any(ok for _, ok in results) else 1
    return 0 if any(ok for _, ok in results) else 1


def gui_run(notes, report, tb, archive_root, result_path):
    """按界面的真实路径跑一遍：建界面 → 填路径 → 点"开始勾稽" → 抽送事件循环。

    和用户实际点击完全同一条链路（工作线程 + Tk 消息泵），用来验证打包版。
    """
    import time

    class SilentBox:
        def showinfo(self, *a, **k):
            pass

        def showerror(self, *a, **k):
            pass

        def showwarning(self, *a, **k):
            pass

        def askyesno(self, *a, **k):
            return False

    globals()['messagebox'] = SilentBox()

    app = AuditCheckApp()
    app.var_notes.set(notes)
    app.var_report.set(report)
    app.var_tb.set(tb)
    app.var_archive.set(archive_root)
    app.var_project.set('')

    state = {'ok': False, 'timeout': False}
    started = [time.time()]

    def finish():
        try:
            text = app.txt_log.get('1.0', 'end')
        except Exception:
            text = ''
        state['ok'] = '[完成]' in text
        try:
            with open(result_path, 'w', encoding='utf-8') as f:
                f.write(text)
        except Exception:
            pass
        try:
            app.quit()
        except Exception:
            pass

    def poll():
        if str(app.btn_run['state']) != 'disabled':
            finish()
            return
        if time.time() - started[0] > 300:
            state['timeout'] = True
            finish()
            return
        app.after(300, poll)

    def kick():
        app.on_run()
        app.after(400, poll)

    app.after(300, kick)
    app.mainloop()
    try:
        app.destroy()
    except Exception:
        pass
    return 0 if state['ok'] else 1


def xls_probe(src, log_path):
    """在打包环境里试 Excel 打开文件，比较中文路径与纯英文路径。"""
    import shutil
    lines = []

    def log(msg):
        lines.append(str(msg))
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines))

    log('TEMP=%s' % os.environ.get('TEMP'))
    log('USERPROFILE=%s' % os.environ.get('USERPROFILE'))
    log('CWD=%s' % os.getcwd())
    log('源文件存在=%s' % os.path.exists(src))

    import win32com.client as win32

    def try_open(path, label):
        xl = None
        try:
            xl = win32.DispatchEx('Excel.Application')
            xl.Visible = False
            xl.DisplayAlerts = False
            wb = xl.Workbooks.Open(os.path.abspath(path), ReadOnly=True)
            log(f'   [{label}] OK 表数={wb.Sheets.Count}')
            wb.Close(False)
            return True
        except Exception as exc:
            log(f'   [{label}] 失败: {exc}')
            return False
        finally:
            if xl is not None:
                try:
                    xl.Quit()
                except Exception:
                    pass

    log('1 原路径（中文）')
    try_open(src, '中文路径')

    ascii_dir = os.path.join(os.path.dirname(os.path.abspath(log_path)), 'ascii')
    os.makedirs(ascii_dir, exist_ok=True)
    ascii_path = os.path.join(ascii_dir, 'probe.xls')
    try:
        shutil.copy2(src, ascii_path)
        log('2 纯英文路径 %s' % ascii_path)
        try_open(ascii_path, '英文路径')
    except Exception as exc:
        log('   复制失败: %s' % (exc,))
    return 0


def word_probe(src, log_path):
    """在打包环境里试 Word 打开文档：原路径 vs 英文暂存路径。"""
    import time
    lines = []

    def log(msg):
        lines.append(str(msg))
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines))

    import com_env
    import win32com.client as win32

    log('TEMP=%s' % os.environ.get('TEMP'))
    staged, copied = com_env.ascii_staging(src, '.doc')
    log('暂存路径=%s 复制=%s' % (staged, copied))

    for label, target in (('原路径', os.path.abspath(src)), ('英文暂存', staged)):
        t0 = time.time()
        word = None
        try:
            with com_env.sanitized():
                word = win32.DispatchEx('Word.Application')
            log(f'   [{label}] DispatchEx {time.time()-t0:.1f}s')
            word.Visible = False
            word.DisplayAlerts = 0
            log(f'   [{label}] 设置属性 {time.time()-t0:.1f}s')
            doc = word.Documents.Open(os.path.abspath(target),
                                      ConfirmConversions=False, ReadOnly=True,
                                      AddToRecentFiles=False)
            log(f'   [{label}] Open OK {time.time()-t0:.1f}s')
            doc.Close(False)
        except Exception as exc:
            log(f'   [{label}] 失败 {time.time()-t0:.1f}s: {exc}')
        finally:
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
    return 0


def main():
    if '--selftest' in sys.argv:
        idx = sys.argv.index('--selftest')
        out = sys.argv[idx + 1] if len(sys.argv) > idx + 1 else 'selftest.txt'
        raise SystemExit(selftest(out))
    if '--wordprobe' in sys.argv:
        idx = sys.argv.index('--wordprobe')
        args = sys.argv[idx + 1:idx + 3]
        if len(args) < 2:
            raise SystemExit(2)
        raise SystemExit(word_probe(*args))
    if '--xlsprobe' in sys.argv:
        idx = sys.argv.index('--xlsprobe')
        args = sys.argv[idx + 1:idx + 3]
        if len(args) < 2:
            raise SystemExit(2)
        raise SystemExit(xls_probe(*args))
    if '--comprobe' in sys.argv:
        idx = sys.argv.index('--comprobe')
        args = sys.argv[idx + 1:idx + 3]
        if len(args) < 2:
            raise SystemExit(2)
        raise SystemExit(com_probe(*args))
    if '--guitest' in sys.argv:
        idx = sys.argv.index('--guitest')
        args = sys.argv[idx + 1:idx + 6]
        if len(args) < 5:
            raise SystemExit(2)
        raise SystemExit(gui_run(*args))
    if '--run' in sys.argv:
        idx = sys.argv.index('--run')
        args = sys.argv[idx + 1:idx + 6]
        if len(args) < 5:
            print('用法：--run <附注> <报表> <TB> <输出目录> <日志文件>')
            raise SystemExit(2)
        raise SystemExit(headless_run(*args))
    AuditCheckApp().mainloop()


if __name__ == '__main__':
    main()
