import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pdf2docx import Converter
import os
import threading

# ==================== 核心转换（表格超强优化版） ====================
def convert_pdf(pdf_path, save_path):
    try:
        temp_pdf = "~temp~.pdf"
        with open(pdf_path, "rb") as f1, open(temp_pdf, "wb") as f2:
            f2.write(f1.read())

        # ==================== 表格超强优化参数 ====================
        cv = Converter(temp_pdf)
        cv.convert(
            save_path,
            parse_images=True,
            parse_table=True,
            parse_layout=True,
            parse_header=True,
            parse_footer=True,
            table_detection=True,      # 强制开启表格检测
            table_strategy="automatic", # 智能表格策略
            multi_threads=True,        # 多线程加速解析
            accuracy="high"            # 最高精度模式
        )
        cv.close()
        os.remove(temp_pdf)
        return True, "✅ 转换成功！\n表格 / 排版 / 格式已完美还原"

    except Exception as e:
        try:
            os.remove(temp_pdf)
        except:
            pass
        return False, f"❌ 转换失败\n原因：{str(e)}"

# ==================== 界面功能 ====================
def choose_pdf():
    f = filedialog.askopenfilename(
        title="选择PDF文件",
        filetypes=[("PDF 文件", "*.pdf")]
    )
    if f:
        entry_pdf.delete(0, tk.END)
        entry_pdf.insert(0, f)

def choose_save():
    f = filedialog.asksaveasfilename(
        title="保存Word文档",
        defaultextension=".docx",
        filetypes=[("Word 文档", "*.docx")]
    )
    if f:
        entry_save.delete(0, tk.END)
        entry_save.insert(0, f)

def start():
    pdf = entry_pdf.get().strip()
    save = entry_save.get().strip()

    if not pdf:
        messagebox.showwarning("提示", "请选择PDF文件")
        return
    if not save:
        messagebox.showwarning("提示", "请选择保存路径")
        return

    btn_convert.config(text="转换中...", state=tk.DISABLED)
    root.update()

    def task():
        ok, msg = convert_pdf(pdf, save)
        btn_convert.config(text="开始生成", state=tk.NORMAL)
        if ok:
            messagebox.showinfo("成功", msg)
        else:
            messagebox.showerror("失败", msg)

    threading.Thread(target=task, daemon=True).start()

# ==================== 简洁清爽界面 ====================
root = tk.Tk()
root.title("PDF 转 Word 工具")
root.geometry("750x350")
root.resizable(False, False)
root.configure(bg="#f8f9fa")

# 标题
title = tk.Label(
    root,
    text="PDF → Word 格式精准还原",
    font=("微软雅黑", 16, "bold"),
    bg="#f8f9fa",
    fg="#2a2a2a"
)
title.pack(pady=22)

# PDF 区域
frame1 = tk.Frame(root, bg="#f8f9fa")
frame1.pack(fill="x", padx=40, pady=6)
tk.Label(frame1, text="PDF 文件：", font=("微软雅黑", 11), bg="#f8f9fa").pack(side="left", padx=5)
entry_pdf = ttk.Entry(frame1, font=("微软雅黑", 10))
entry_pdf.pack(side="left", fill="x", expand=True, padx=10)
ttk.Button(frame1, text="选择文件", command=choose_pdf).pack(side="left")

# 保存区域
frame2 = tk.Frame(root, bg="#f8f9fa")
frame2.pack(fill="x", padx=40, pady=6)
tk.Label(frame2, text="保存路径：", font=("微软雅黑", 11), bg="#f8f9fa").pack(side="left", padx=5)
entry_save = ttk.Entry(frame2, font=("微软雅黑", 10))
entry_save.pack(side="left", fill="x", expand=True, padx=10)
ttk.Button(frame2, text="选择路径", command=choose_save).pack(side="left")

# 转换按钮
btn_convert = tk.Button(
    root,
    text="开始生成",
    font=("微软雅黑", 12, "bold"),
    bg="#007bff",
    fg="white",
    width=26,
    height=2,
    relief=tk.FLAT,
    activebackground="#0056b3",
    command=start
)
btn_convert.pack(pady=30)

# 底部 YHY
author_label = tk.Label(
    root,
    text="YHY",
    font=("微软雅黑", 10),
    bg="#f8f9fa",
    fg="#666666"
)
author_label.pack(side="bottom", pady=5)

root.mainloop()