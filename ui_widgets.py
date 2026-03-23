# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk

class LabeledEntry(ttk.Frame):
    def __init__(self, master, label: str, width: int=30):
        super().__init__(master)
        self.var = tk.StringVar()
        ttk.Label(self, text=label).grid(row=0, column=0, sticky="w", padx=4, pady=2)
        self.entry = ttk.Entry(self, textvariable=self.var, width=width)
        self.entry.grid(row=0, column=1, sticky="we", padx=4, pady=2)
        self.columnconfigure(1, weight=1)

class LabeledCombo(ttk.Frame):
    def __init__(self, master, label: str, values, width: int=28):
        super().__init__(master)
        self.var = tk.StringVar()
        ttk.Label(self, text=label).grid(row=0, column=0, sticky="w", padx=4, pady=2)
        self.combo = ttk.Combobox(self, textvariable=self.var, values=list(values), width=width, state="readonly")
        self.combo.grid(row=0, column=1, sticky="we", padx=4, pady=2)
        self.columnconfigure(1, weight=1)

def make_tree(master, columns, widths=None):
    tree = ttk.Treeview(master, columns=columns, show="headings")
    for i, col in enumerate(columns):
        tree.heading(col, text=col)
        w = widths[i] if widths else 140
        tree.column(col, width=w, anchor="w")
    vsb = ttk.Scrollbar(master, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=vsb.set)
    return tree, vsb

def show_info(parent, title, msg):
    win = tk.Toplevel(parent)
    win.title(title)
    win.geometry("520x260")
    txt = tk.Text(win, wrap="word")
    txt.insert("1.0", msg)
    txt.configure(state="disabled")
    txt.pack(fill="both", expand=True, padx=8, pady=8)
    ttk.Button(win, text="Закрыть", command=win.destroy).pack(pady=6)
