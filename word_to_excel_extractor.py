# -*- coding: utf-8 -*-
"""
Word 审计报告附注自动提取工具 (Word-to-Excel Extractor)
功能：
从审计报告 Word 文档 (.doc / .docx) 中自动定位《六、财务报表主要项目附注》，
将该段落下所有小标题及报表表格完整抽取并写入 Excel 工作表，生成待勾稽的基础 Excel 文件。
"""

import os
import re
import openpyxl
import win32com.client

def extract_notes_from_word(doc_filename=None, out_excel='extracted_notes.xlsx'):
    # Locate Word file if not provided
    if not doc_filename:
        docs = [f for f in os.listdir('.') if f.endswith(('.doc', '.docx')) and not f.startswith('~$')]
        if not docs:
            raise FileNotFoundError("未在当前目录下找到 Word 报告文件 (.doc / .docx)")
        doc_filename = docs[0]

    abs_doc = os.path.abspath(doc_filename)
    print(f"正在从 Word 文件提取附注内容: {doc_filename} ...")

    word = win32com.client.gencache.EnsureDispatch('Word.Application') if hasattr(win32com.client, 'gencache') else win32com.client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = False

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    row_idx = 1
    try:
        doc = word.Documents.Open(FileName=abs_doc, ConfirmConversions=False, ReadOnly=True)
        in_target_section = False
        
        # Traverse paragraphs and tables in order
        for p in doc.Paragraphs:
            text = p.Range.Text.strip().replace('\r', '').replace('\x07', '')
            if not text:
                continue
            
            # Start of target section: "六、财务报表主要项目附注" or similar
            if re.match(r'^[六七五四三二一\d]+[、\.]\s*财务报表主要项目附注', text):
                in_target_section = True
                ws.cell(row_idx, 1, text)
                row_idx += 1
                continue
            
            # End of target section: next level 1 heading like "七、"
            if in_target_section and re.match(r'^[七八九十\d]+[、\.]\s*(关联方|承诺及或有事项|资产负债表日后事项|其他重要事项)', text):
                print(f"检测到附注结束标识: {text}")
                break

            if in_target_section:
                # Check if paragraph contains table or standalone title
                if p.Range.Tables.Count == 0:
                    ws.cell(row_idx, 1, text)
                    row_idx += 1

        doc.Close(False)
    finally:
        word.Quit()

    print(f"提取完成，共写入 {row_idx} 行。保存至 {out_excel} ...")
    wb.save(out_excel)
    return out_excel

if __name__ == '__main__':
    import sys
    target_dir = '.'
    if len(sys.argv) > 1:
        target_dir = sys.argv[1]
    else:
        docs_curr = [f for f in os.listdir('.') if f.endswith(('.doc', '.docx')) and not f.startswith('~$')]
        if not docs_curr:
            subdirs = [d for d in os.listdir('.') if os.path.isdir(d) and not d.startswith('.') and not d.startswith('__')]
            valid_subdirs = [d for d in subdirs if any(f.endswith(('.doc', '.docx')) and not f.startswith('~$') for f in os.listdir(d))]
            if len(valid_subdirs) == 1:
                target_dir = valid_subdirs[0]
                print(f"自动定位到公司子目录: {target_dir}")
            elif len(valid_subdirs) > 1:
                print(f"检测到多个公司目录: {valid_subdirs}")
                print(f"请指定公司目录运行，例如: python word_to_excel_extractor.py \"{valid_subdirs[0]}\"")
                sys.exit(0)
    
    docs = [f for f in os.listdir(target_dir) if f.endswith(('.doc', '.docx')) and not f.startswith('~$')]
    if docs:
        doc_path = os.path.join(target_dir, docs[0])
        out_path = os.path.join(target_dir, 'check.xlsx')
        extract_notes_from_word(doc_path, out_path)
    else:
        print(f"在目录 {target_dir} 中未找到 Word 文档")
