#!/usr/bin/env python
# -*- coding:utf-8 -*-
# author:HeWenjun
# datetime:2023/3/10 15:28
# software: PyCharm
import tkinter as tk
from tkinter import filedialog
from AICode import KeNuo

root = tk.Tk()
root.title("My GUI")
root.geometry("300x500+{}+{}".format(int(root.winfo_screenwidth() / 2 - 200), int(root.winfo_screenheight() / 2 - 150)))


def upload_file():
    filename = filedialog.askopenfilename(initialdir="/", title="Select file",
                                          filetypes=(("Excel files", "*.xls;*.xlsx"), ("all files", "*.*")))
    text.insert(tk.END, filename + '\n')


def run_excel():
    file_path = text.get('1.0', tk.END).strip()
    print("正在处理文件：{}".format(file_path))
    KeNuo(file_path).data_writer()
    label3.config(text="已完成")


# 创建Label控件
label = tk.Label(root, text="点击选择文件", width=35, height=10, bg="white")
label.place(relx=0.5, rely=0.4, anchor="center")
label.bind("<Button-1>", lambda e: upload_file())
# 创建展示文件名的多行文本框
text = tk.Text(root, width=35, height=5, bg="white")
text.place(relx=0.5, rely=0.7, anchor="center")
# 创建运行按钮
button = tk.Button(root, text="运行", width=10, height=2, command=run_excel)
button.place(relx=0.5, rely=0.85, anchor="center")
# 创建提示信息
label3 = tk.Label(root, text="", width=10, height=2, bg="white")
label3.place(relx=0.5, rely=0.95, anchor="center")
root.mainloop()
