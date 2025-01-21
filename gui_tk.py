# gui_tk.py

import tkinter as tk
from tkinter import filedialog, messagebox
import os

from core import (
    read_main_article,
    generate_text_and_letter,
    merge_text_and_letter,
    clean_temp_files
)

def run_gui():
    root = tk.Tk()
    root.title('台灣郵局存證信函對齊器 - 簡易GUI')

    # 文章路徑
    def browse_article():
        path = filedialog.askopenfilename(filetypes=[('Text Files', '*.txt'), ('All Files', '*.*')])
        if path:
            article_path_var.set(path)

    article_label = tk.Label(root, text='存證信函文字檔路徑')
    article_label.grid(row=0, column=0, sticky='e', padx=5, pady=5)

    article_path_var = tk.StringVar()
    article_entry = tk.Entry(root, textvariable=article_path_var, width=40)
    article_entry.grid(row=0, column=1, padx=5, pady=5)
    article_browse_btn = tk.Button(root, text='選擇檔案', command=browse_article)
    article_browse_btn.grid(row=0, column=2, padx=5, pady=5)

    # 寄件人
    sender_name_var = tk.StringVar()
    sender_addr_var = tk.StringVar()

    sender_name_label = tk.Label(root, text='寄件人姓名')
    sender_name_label.grid(row=1, column=0, sticky='e', padx=5, pady=5)
    sender_name_entry = tk.Entry(root, textvariable=sender_name_var, width=40)
    sender_name_entry.grid(row=1, column=1, padx=5, pady=5)

    sender_addr_label = tk.Label(root, text='寄件人地址')
    sender_addr_label.grid(row=2, column=0, sticky='e', padx=5, pady=5)
    sender_addr_entry = tk.Entry(root, textvariable=sender_addr_var, width=40)
    sender_addr_entry.grid(row=2, column=1, padx=5, pady=5)

    # 收件人
    receiver_name_var = tk.StringVar()
    receiver_addr_var = tk.StringVar()

    receiver_name_label = tk.Label(root, text='收件人姓名')
    receiver_name_label.grid(row=3, column=0, sticky='e', padx=5, pady=5)
    receiver_name_entry = tk.Entry(root, textvariable=receiver_name_var, width=40)
    receiver_name_entry.grid(row=3, column=1, padx=5, pady=5)

    receiver_addr_label = tk.Label(root, text='收件人地址')
    receiver_addr_label.grid(row=4, column=0, sticky='e', padx=5, pady=5)
    receiver_addr_entry = tk.Entry(root, textvariable=receiver_addr_var, width=40)
    receiver_addr_entry.grid(row=4, column=1, padx=5, pady=5)

    # 輸出檔名
    output_filename_var = tk.StringVar(value='output.pdf')
    output_filename_label = tk.Label(root, text='輸出檔名')
    output_filename_label.grid(row=5, column=0, sticky='e', padx=5, pady=5)
    output_filename_entry = tk.Entry(root, textvariable=output_filename_var, width=40)
    output_filename_entry.grid(row=5, column=1, padx=5, pady=5)

    def generate_pdf():
        article_file = article_path_var.get().strip()
        if not article_file or not os.path.exists(article_file):
            messagebox.showerror('錯誤', '存證信函文字檔路徑不存在')
            return
        main_text = read_main_article(article_file)

        s_name = sender_name_var.get().strip()
        s_addr = sender_addr_var.get().strip()
        r_name = receiver_name_var.get().strip()
        r_addr = receiver_addr_var.get().strip()

        out_file = output_filename_var.get().strip()
        if not out_file.endswith('.pdf'):
            out_file += '.pdf'

        # 以舊版參數結構對應
        senders = [[s_name]] if s_name else []
        senders_addr = [s_addr] if s_addr else []
        receivers = [[r_name]] if r_name else []
        receivers_addr = [r_addr] if r_addr else []

        text_path, letter_path = generate_text_and_letter(
            senders=senders,
            senders_addr=senders_addr,
            receivers=receivers,
            receivers_addr=receivers_addr,
            ccs=[],   # 這裡簡化，不填副本
            cc_addr=[],
            main_text=main_text
        )

        merge_text_and_letter(text_path, letter_path, out_file)
        clean_temp_files(text_path, letter_path)

        messagebox.showinfo('完成', f'PDF 已產生：{out_file}')

    generate_btn = tk.Button(root, text='產生 PDF', command=generate_pdf)
    generate_btn.grid(row=6, column=1, padx=5, pady=15)

    root.mainloop()
