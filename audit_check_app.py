# -*- coding: utf-8 -*-
"""
财务报表附注自动化勾稽系统 - 桌面交互与一键归档应用
包含功能：
1. 文件提交与项目自动识别（报告附注、财务报表、TB试算平衡表）
2. Word 附注主要项目编号诊断与一键规范重编号
3. 报表 Excel 附注编号一键修复与补全
4. 自动化全面勾稽引擎（附注内表加减、报表联动勾稽、TB现金流逐格勾稽）
5. 自动多版本递增与归档（附注check_项目名_v1.xlsx、v2...）
"""

import os
import sys
import re
import shutil
import threading
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# 导入核心管理器与引擎
from note_number_manager import (
    extract_word_note_items,
    diagnose_note_continuity,
    renumber_word_file,
    fix_financial_statements_note_numbers
)

GITHUB_URL = "https://github.com/zhanghongruiszh/financial-notes-audit-tool"
AUTHOR_EMAIL = "henryzhang825@gmail.com"

def infer_project_name(*file_paths):
    """
    从提交的文件名中智能提取公司/项目名。
    若无法提取，返回空字符串，绝不预填任何无关公司名。
    """
    for p in file_paths:
        if not p:
            continue
        base = os.path.basename(p)
        clean = re.sub(r'^(TB|tb|1|2|3|4|5|6|7|8|9|0|[·\-_\s])+', '', base)
        clean = re.sub(r'[\-_]?(三年一期|申报|审定|合并|单体|资产负债表|利润表|现金流量表|财务报表附注|财务报表|主要项目附注|附注).*$', '', clean, flags=re.IGNORECASE)
        clean = re.sub(r'[\(（].*?[\)）]', '', clean)
        clean = clean.strip('-_ ')
        if len(clean) >= 2:
            return clean
    return ""

class AuditCheckApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("财务报表附注自动化勾稽系统 v2.0")
        self.geometry("940x700")
        self.minsize(880, 620)
        
        self.center_window()
        
        self.style = ttk.Style()
        try:
            self.style.theme_use('vista')
        except:
            pass
            
        self.create_widgets()
        
    def center_window(self):
        self.update_idletasks()
        w = 940
        h = 700
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2) - 20
        self.geometry(f"{w}x{h}+{x}+{y}")
        
    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="15 12 15 12")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 1. 顶部标题栏
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 6))
        
        lbl_title = tk.Label(
            header_frame, 
            text="财务报表附注自动化勾稽与编号治理系统", 
            font=("微软雅黑", 14, "bold"),
            fg="#1F497D"
        )
        lbl_title.pack(anchor=tk.W)
        
        lbl_desc = tk.Label(
            header_frame,
            text="功能模块：Word附注编号拓扑自检与顺序重排 | 报表附注号一键修复补全 | 附注内表加减与TB现金流逐格勾稽",
            font=("微软雅黑", 9),
            fg="#555555"
        )
        lbl_desc.pack(anchor=tk.W, pady=(2, 4))
        
        # 项目链接与联系邮箱条
        info_bar = ttk.Frame(header_frame)
        info_bar.pack(fill=tk.X, pady=(0, 4))
        
        lbl_git_prefix = tk.Label(info_bar, text="开源项目地址：", font=("微软雅黑", 8), fg="#333333")
        lbl_git_prefix.pack(side=tk.LEFT)
        
        lbl_git_link = tk.Label(
            info_bar, 
            text=GITHUB_URL, 
            font=("微软雅黑", 8, "underline"), 
            fg="#0969DA", 
            cursor="hand2"
        )
        lbl_git_link.pack(side=tk.LEFT)
        lbl_git_link.bind("<Button-1>", lambda e: webbrowser.open_new_tab(GITHUB_URL))
        
        lbl_star = tk.Label(info_bar, text=" (欢迎点击访问并 Star 支持)", font=("微软雅黑", 8), fg="#666666")
        lbl_star.pack(side=tk.LEFT)
        
        lbl_email = tk.Label(
            info_bar, 
            text=f"问题反馈邮箱：{AUTHOR_EMAIL}", 
            font=("微软雅黑", 8), 
            fg="#1F497D"
        )
        lbl_email.pack(side=tk.RIGHT)
        
        # 2. 文件选择区（初始全部置空，不残留任何历史文件名）
        grp_files = ttk.LabelFrame(main_frame, text=" 提交原始文件 ", padding="12 10 12 10")
        grp_files.pack(fill=tk.X, pady=(0, 10))
        
        # 报告附注
        ttk.Label(grp_files, text="1. 报告附注:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.var_notes = tk.StringVar(value="")
        self.ent_notes = ttk.Entry(grp_files, textvariable=self.var_notes, width=64)
        self.ent_notes.grid(row=0, column=1, sticky=tk.EW, padx=8, pady=5)
        btn_notes = ttk.Button(grp_files, text="选择附注...", command=self.browse_notes)
        btn_notes.grid(row=0, column=2, sticky=tk.W, pady=5)
        
        # 财务报表
        ttk.Label(grp_files, text="2. 财务报表:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.var_report = tk.StringVar(value="")
        self.ent_report = ttk.Entry(grp_files, textvariable=self.var_report, width=64)
        self.ent_report.grid(row=1, column=1, sticky=tk.EW, padx=8, pady=5)
        btn_report = ttk.Button(grp_files, text="选择报表...", command=self.browse_report)
        btn_report.grid(row=1, column=2, sticky=tk.W, pady=5)
        
        # TB 试算平衡表
        ttk.Label(grp_files, text="3. TB试算表:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.var_tb = tk.StringVar(value="")
        self.ent_tb = ttk.Entry(grp_files, textvariable=self.var_tb, width=64)
        self.ent_tb.grid(row=2, column=1, sticky=tk.EW, padx=8, pady=5)
        btn_tb = ttk.Button(grp_files, text="选择TB表...", command=self.browse_tb)
        btn_tb.grid(row=2, column=2, sticky=tk.W, pady=5)
        
        grp_files.columnconfigure(1, weight=1)
        
        # 3. 项目与归档配置区
        grp_cfg = ttk.LabelFrame(main_frame, text=" 归档与版本控制 ", padding="12 8 12 8")
        grp_cfg.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(grp_cfg, text="识别项目名:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.var_project = tk.StringVar(value="")
        self.ent_project = ttk.Entry(grp_cfg, textvariable=self.var_project, width=22)
        self.ent_project.grid(row=0, column=1, sticky=tk.W, padx=8, pady=5)
        
        ttk.Label(grp_cfg, text="成果命名规则:").grid(row=0, column=2, sticky=tk.W, pady=5, padx=(15, 0))
        lbl_rule = ttk.Label(grp_cfg, text="附注check_项目名_v1.xlsx (版本号递增保留)", foreground="#1F497D", font=("微软雅黑", 8, "bold"))
        lbl_rule.grid(row=0, column=3, sticky=tk.W, pady=5)
        
        ttk.Label(grp_cfg, text="归档存放目录:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.var_archive = tk.StringVar(value=os.path.abspath('.'))
        self.ent_archive = ttk.Entry(grp_cfg, textvariable=self.var_archive, width=54)
        self.ent_archive.grid(row=1, column=1, columnspan=2, sticky=tk.EW, padx=8, pady=5)
        btn_archive = ttk.Button(grp_cfg, text="选择目录...", command=self.browse_archive)
        btn_archive.grid(row=1, column=3, sticky=tk.W, pady=5)
        
        grp_cfg.columnconfigure(1, weight=1)
        
        # 4. 操作按钮区
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 主按钮
        self.btn_run = tk.Button(
            btn_frame, 
            text="开始自动化全面勾稽 (生成Check表)", 
            font=("微软雅黑", 10, "bold"),
            bg="#1F497D", fg="white", 
            activebackground="#153358", activeforeground="white",
            relief=tk.RAISED, cursor="hand2",
            padx=14, pady=7,
            command=self.on_start_clicked
        )
        self.btn_run.pack(side=tk.LEFT, padx=(0, 10))
        
        # 辅助功能按钮
        self.btn_word_fix = tk.Button(
            btn_frame,
            text="Word附注编号诊断与重排",
            font=("微软雅黑", 9),
            bg="#EBF1F5", fg="#1F497D",
            cursor="hand2", padx=10, pady=6,
            command=self.on_word_fix_clicked
        )
        self.btn_word_fix.pack(side=tk.LEFT, padx=(0, 10))
        
        self.btn_stmt_fix = tk.Button(
            btn_frame,
            text="报表附注号一键修复与补全",
            font=("微软雅黑", 9),
            bg="#EBF1F5", fg="#1F497D",
            cursor="hand2", padx=10, pady=6,
            command=self.on_stmt_fix_clicked
        )
        self.btn_stmt_fix.pack(side=tk.LEFT, padx=(0, 10))
        
        self.btn_open_folder = tk.Button(
            btn_frame,
            text="打开项目归档文件夹",
            font=("微软雅黑", 9),
            bg="#F2F2F2", fg="#333333",
            cursor="hand2", padx=10, pady=6,
            command=self.open_archive_folder
        )
        self.btn_open_folder.pack(side=tk.RIGHT)
        
        # 5. 执行日志与反馈区
        grp_log = ttk.LabelFrame(main_frame, text=" 执行日志与诊断反馈 ", padding="10 8 10 8")
        grp_log.pack(fill=tk.BOTH, expand=True)
        
        self.txt_log = tk.Text(grp_log, wrap=tk.WORD, font=("Consolas", 9), bg="#F8F9FA", fg="#262626")
        self.txt_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(grp_log, orient=tk.VERTICAL, command=self.txt_log.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.txt_log.config(yscrollcommand=scrollbar.set)
        
        self.log("[系统状态] 就绪。请点击上方按钮选择需核对的报告附注、财务报表及TB底稿。")
        
    def log(self, msg):
        self.txt_log.insert(tk.END, msg + "\n")
        self.txt_log.see(tk.END)
        self.update_idletasks()
        
    def browse_notes(self):
        path = filedialog.askopenfilename(
            title="选择报告附注文件",
            filetypes=[("Word/Excel 附注文件", "*.doc *.docx *.xlsx"), ("所有文件", "*.*")]
        )
        if path:
            self.var_notes.set(path)
            self.auto_guess_project()
            
    def browse_report(self):
        path = filedialog.askopenfilename(
            title="选择财务报表文件",
            filetypes=[("Excel 财务报表", "*.xlsx"), ("所有文件", "*.*")]
        )
        if path:
            self.var_report.set(path)
            self.auto_guess_project()
            
    def browse_tb(self):
        path = filedialog.askopenfilename(
            title="选择 TB 试算平衡表底稿",
            filetypes=[("Excel 工作底稿", "*.xls *.xlsx"), ("所有文件", "*.*")]
        )
        if path:
            self.var_tb.set(path)
            self.auto_guess_project()
            
    def browse_archive(self):
        path = filedialog.askdirectory(title="选择归档存放目录")
        if path:
            self.var_archive.set(path)
            
    def auto_guess_project(self):
        if not self.var_project.get().strip():
            inferred = infer_project_name(self.var_report.get(), self.var_notes.get(), self.var_tb.get())
            if inferred:
                self.var_project.set(inferred)
                
    def on_word_fix_clicked(self):
        """Word附注编号拓扑自检与顺序重排"""
        doc_file = self.var_notes.get().strip()
        if not doc_file or not os.path.exists(doc_file) or not doc_file.endswith(('.doc', '.docx')):
            messagebox.showwarning("提示", "请先在上方【报告附注】处选择有效的 Word 附注文件 (.doc / .docx)。")
            return
            
        self.log(f"\n[处理] 正在对 Word 附注编号进行健康度检测: {os.path.basename(doc_file)} ...")
        
        def run():
            try:
                c_pref, p_pref, c_items, p_items = extract_word_note_items(doc_file)
                self.log(f"[信息] 大标题章节：合并附注 [{c_pref}]，母公司附注 [{p_pref}]")
                self.log(f"[信息] 共提取到合并附注子科目 {len(c_items)} 项，母公司附注子科目 {len(p_items)} 项。")
                
                is_ok_c, issues_c = diagnose_note_continuity(c_items, "合并附注")
                is_ok_p, issues_p = diagnose_note_continuity(p_items, "母公司附注")
                
                all_issues = issues_c + (issues_p if p_items else [])
                if not all_issues:
                    self.log("[检查结果] Word 附注子项目编号完全连续递增（1..N），未发现跳号或重号。")
                    messagebox.showinfo("诊断结果", "Word 附注编号完全连续，无需重排。")
                else:
                    self.log(f"[警告] 发现编号异常共 {len(all_issues)} 处：")
                    for iss in all_issues[:15]:
                        self.log(f"  - {iss}")
                    if len(all_issues) > 15:
                        self.log(f"  ... 另有 {len(all_issues) - 15} 处已省略。")
                        
                    ans = messagebox.askyesno(
                        "发现编号异常",
                        f"检测到附注编号存在 {len(all_issues)} 处异常（断号/重号/未编号）。\n\n是否立即执行【一键顺序规范重编号】？\n系统将生成编号已规范的副本并保留原文档。"
                    )
                    if ans:
                        self.log("[处理] 正在执行 Word 附注主要项目一键顺序重编号 ...")
                        out_doc, changes = renumber_word_file(doc_file)
                        self.log(f"[成功] 重编号完成，新文件保存在：{out_doc}")
                        self.log(f"[信息] 共规范重排 {len(changes)} 处子项目编号。")
                        messagebox.showinfo("完成", f"已成功规范重排附注子项目编号。\n文件保存在：\n{out_doc}")
            except Exception as e:
                self.log(f"[错误] 处理失败: {str(e)}")
                messagebox.showerror("错误", f"Word 附注处理出错: {str(e)}")
                
        threading.Thread(target=run, daemon=True).start()
        
    def on_stmt_fix_clicked(self):
        """报表 Excel 附注编号一键修复补全"""
        notes_file = self.var_notes.get().strip()
        report_file = self.var_report.get().strip()
        
        if not notes_file or not os.path.exists(notes_file):
            messagebox.showwarning("提示", "请先选择有效的报告附注文件。")
            return
        if not report_file or not os.path.exists(report_file):
            messagebox.showwarning("提示", "请先选择有效的财务报表 Excel 文件。")
            return
            
        self.log(f"\n[处理] 正在根据附注一键修复与补全报表附注编号 ...")
        
        def run():
            try:
                c_pref, p_pref, c_items, p_items = extract_word_note_items(notes_file)
                items_map = {item[2]: item[1] for item in c_items if item[1] is not None}
                
                out_excel, changes = fix_financial_statements_note_numbers(
                    report_file, items_map, chapter_prefix=c_pref
                )
                self.log(f"[成功] 报表附注编号处理完成，已生成新文件：{out_excel}")
                self.log(f"[信息] 共补全/修正 {len(changes)} 处附注编号：")
                for ch in changes[:20]:
                    self.log(f"  - {ch}")
                if len(changes) > 20:
                    self.log(f"  ... 另有 {len(changes) - 20} 处已省略。")
                messagebox.showinfo("完成", f"报表附注编号已成功修复与补全。\n共处理 {len(changes)} 处。\n文件已保存至：\n{out_excel}")
            except Exception as e:
                self.log(f"[错误] 报表修复失败: {str(e)}")
                messagebox.showerror("错误", f"报表附注编号修复出错: {str(e)}")
                
        threading.Thread(target=run, daemon=True).start()
        
    def on_start_clicked(self):
        notes_file = self.var_notes.get().strip()
        report_file = self.var_report.get().strip()
        tb_file = self.var_tb.get().strip()
        project_name = self.var_project.get().strip()
        base_archive = self.var_archive.get().strip()
        
        if not notes_file or not os.path.exists(notes_file):
            messagebox.showwarning("提示", "请选择有效的报告附注文件。")
            return
        if not report_file or not os.path.exists(report_file):
            messagebox.showwarning("提示", "请选择有效的财务报表文件。")
            return
        if not tb_file or not os.path.exists(tb_file):
            messagebox.showwarning("提示", "请选择有效的 TB 试算平衡表。")
            return
            
        if not project_name:
            project_name = infer_project_name(report_file, notes_file, tb_file)
            if not project_name:
                project_name = "未命名项目"
            self.var_project.set(project_name)
            
        company_dir = os.path.join(base_archive, project_name)
        os.makedirs(company_dir, exist_ok=True)
        self.last_company_dir = company_dir
        
        self.btn_run.config(state=tk.DISABLED, text="勾稽运算中...")
        self.log(f"\n=======================================================")
        self.log(f"开始执行自动化全面勾稽：项目【{project_name}】")
        self.log(f"归档目录: {company_dir}")
        self.log(f"=======================================================")
        
        threading.Thread(target=self.run_full_check, args=(notes_file, report_file, tb_file, project_name, company_dir), daemon=True).start()
        
    def run_full_check(self, notes_file, report_file, tb_file, project_name, company_dir):
        try:
            # 1. 归档原始文件
            self.log("[步骤 1/4] 归档原始提交文件到项目专用文件夹 ...")
            def safe_copy(src, dst_dir):
                if not os.path.exists(src):
                    return None
                tgt = os.path.join(dst_dir, os.path.basename(src))
                if os.path.abspath(src) != os.path.abspath(tgt):
                    shutil.copy2(src, tgt)
                return tgt
                
            arc_notes = safe_copy(notes_file, company_dir)
            arc_report = safe_copy(report_file, company_dir)
            arc_tb = safe_copy(tb_file, company_dir)
            
            # 2. 如果附注是 Word，且无 check.xlsx，则提取
            check_xlsx_path = os.path.join(company_dir, "check.xlsx")
            if not os.path.exists(check_xlsx_path) and arc_notes and arc_notes.endswith(('.doc', '.docx')):
                self.log("[步骤 2/4] 从 Word 附注中抽取表格内容至 check.xlsx ...")
                from word_to_excel_extractor import extract_notes_from_word
                extract_notes_from_word(arc_notes, out_excel=check_xlsx_path)
            elif arc_notes and arc_notes.endswith('.xlsx') and os.path.basename(arc_notes) != "check.xlsx":
                shutil.copy2(arc_notes, check_xlsx_path)
                
            # 3. 运行勾稽引擎
            self.log("[步骤 3/4] 运行核心勾稽引擎（内表加减、报表联动、TB现金流核对）...")
            import audit_check_engine
            proj, dst_file, ver_num = audit_check_engine.build_full_system(target_dir=company_dir)
            
            self.log("[步骤 4/4] 勾稽完成。")
            self.log(f"[成果文件] {dst_file}")
            self.log(f"[版本号] v{ver_num}")
            self.log(f"[归档目录] {company_dir}")
            
            messagebox.showinfo(
                "勾稽完成",
                f"项目【{project_name}】附注勾稽已完成。\n\n成果文件：\n{os.path.basename(dst_file)}\n版本：v{ver_num}\n\n已归档至：\n{company_dir}"
            )
        except Exception as e:
            self.log(f"[错误] 运行过程中发生错误: {str(e)}")
            messagebox.showerror("运行错误", f"勾稽处理失败:\n{str(e)}")
        finally:
            self.btn_run.config(state=tk.NORMAL, text="开始自动化全面勾稽 (生成Check表)")
            
    def open_archive_folder(self):
        folder = getattr(self, 'last_company_dir', None)
        if not folder or not os.path.exists(folder):
            proj = self.var_project.get().strip()
            if proj:
                folder = os.path.join(self.var_archive.get().strip(), proj)
        if folder and os.path.exists(folder):
            os.startfile(folder)
        else:
            messagebox.showinfo("提示", "归档目录尚未创建，请先执行勾稽。")

if __name__ == "__main__":
    app = AuditCheckApp()
    app.mainloop()
