#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
股票私人助手 PRO GUI版（同花顺导入版 UTF-8）
功能：筛选 + 显示代码与名称 + 导出同花顺可导入格式
"""

import pandas as pd
import datetime
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

def clean_column_name(name):
    name = re.sub(r"[\[\(（【].*?[\]\)）】]", "", str(name))
    return name.strip()

def parse_amount(s):
    s = str(s).strip().replace(",", "").replace("+", "").replace(" ", "")
    try:
        if "亿" in s:
            return float(s.replace("亿", "")) * 1e8
        if "万" in s:
            return float(s.replace("万", "")) * 1e4
        return float(s)
    except:
        return None

def parse_free_float(s):
    s = str(s).strip().replace(",", "").replace("+", "").replace(" ", "")
    try:
        if "亿" in s:
            return float(s.replace("亿", ""))
        if "万" in s:
            return float(s.replace("万", "")) / 10000.0
        num = float(s)
        return num / 1e8 if num > 1e7 else num
    except:
        return None

def fuzzy_get(df, keywords):
    for col in df.columns:
        for k in keywords:
            if k in col:
                return col
    return None

def run_model(df, model_type="A"):
    df = df.copy()
    df = df[~df["代码"].astype(str).str.startswith("SH688")]
    df = df[~df["名称"].astype(str).str.contains("ST", na=False)]
    df["涨幅_float"] = pd.to_numeric(df["涨幅"].astype(str).str.replace("%", "").str.replace(" ", ""), errors="coerce")
    df = df[df["涨幅_float"] < 9.5]
    df["散户_float"] = pd.to_numeric(df["散户数量"].astype(str).str.replace(" ", ""), errors="coerce")
    df = df[df["散户_float"] <= -20]
    df["自由流通_float"] = df["自由流通"].apply(parse_free_float)
    df = df[df["自由流通_float"] < 3.5]
    df["主力净额_float"] = df["主力净额"].apply(parse_amount)
    df["现价_float"] = pd.to_numeric(df["现价"].astype(str).str.replace(" ", ""), errors="coerce")
    df["振幅_float"] = pd.to_numeric(df["振幅"].astype(str).str.replace("%", "").str.replace(" ", ""), errors="coerce")
    df["换手_float"] = pd.to_numeric(df["换手"].astype(str).str.replace("%", "").str.replace(" ", ""), errors="coerce")
    df["量比_float"] = pd.to_numeric(df["量比"].astype(str).str.replace(" ", ""), errors="coerce")

    if model_type == "A":
        cond = (
            (df["现价_float"] < 20) &
            (df["主力净额_float"] >= 8e6) &
            (df["振幅_float"] <= 10) &
            (df["换手_float"].between(6, 13)) &
            (df["量比_float"].between(1.1, 2.0))
        )
    else:
        cond = (
            (df["现价_float"] < 20) &
            (df["主力净额_float"] >= 1e7) &
            (df["振幅_float"] <= 14) &
            (df["换手_float"].between(9, 22)) &
            (df["量比_float"].between(1.2, 2.8))
        )

    return df[cond][["代码", "名称"]]

class StockAssistantApp:
    def __init__(self, root):
        self.root = root
        self.root.title("股票私人助手 PRO（同花顺导入版 UTF-8）")
        self.df = None
        self.result_A = None
        self.result_B = None
        self.build_ui()

    def build_ui(self):
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill="x")

        ttk.Label(frame, text="选择股票数据文件（CSV/Excel）:").pack(side="left")
        self.file_label = ttk.Label(frame, text="未选择文件", width=40)
        self.file_label.pack(side="left", padx=5)
        ttk.Button(frame, text="浏览", command=self.load_file).pack(side="left")

        btn_frame = ttk.Frame(self.root, padding=10)
        btn_frame.pack(fill="x")
        ttk.Button(btn_frame, text="按模型 A 筛选", command=lambda: self.run_model_gui("A")).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="按模型 B 筛选", command=lambda: self.run_model_gui("B")).pack(side="left", padx=10)

        self.tree = ttk.Treeview(self.root, columns=("代码", "名称"), show="headings", height=15)
        self.tree.heading("代码", text="代码")
        self.tree.heading("名称", text="名称")
        self.tree.pack(fill="both", expand=True, padx=10, pady=5)

        ttk.Button(self.root, text="导出同花顺格式 TXT 文件", command=self.export_txt).pack(pady=5)

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xls;*.xlsx")])
        if not path:
            return
        self.file_label.config(text=Path(path).name)

        if path.endswith(".csv"):
            df = pd.read_csv(path, encoding="utf-8", dtype=str)
        else:
            df = pd.read_excel(path, dtype=str)
        df.columns = [clean_column_name(c) for c in df.columns]

        mapping = {
            "代码": fuzzy_get(df, ["代码"]),
            "名称": fuzzy_get(df, ["名称"]),
            "现价": fuzzy_get(df, ["现价", "价格"]),
            "涨幅": fuzzy_get(df, ["涨幅"]),
            "振幅": fuzzy_get(df, ["振幅", "振"]),
            "散户数量": fuzzy_get(df, ["散户"]),
            "主力净额": fuzzy_get(df, ["主力净额", "主力净"]),
            "换手": fuzzy_get(df, ["换手"]),
            "自由流通": fuzzy_get(df, ["自由流通", "流通股"]),
            "量比": fuzzy_get(df, ["量比"])
        }
        df = df.rename(columns={v: k for k, v in mapping.items() if v})
        self.df = df
        messagebox.showinfo("成功", "文件加载成功！")

    def run_model_gui(self, model_type):
        if self.df is None:
            messagebox.showwarning("提示", "请先选择数据文件！")
            return

        result = run_model(self.df, model_type)
        for i in self.tree.get_children():
            self.tree.delete(i)
        for _, row in result.iterrows():
            self.tree.insert("", "end", values=(row["代码"], row["名称"]))

        if model_type == "A":
            self.result_A = result
        else:
            self.result_B = result

        messagebox.showinfo("完成", f"模型 {model_type} 筛选完成，共 {len(result)} 支股票。")

    def export_txt(self):
        date_str = datetime.datetime.now().strftime("%Y%m%d")

        if self.result_A is not None and not self.result_A.empty:
            outA = Path(f"{date_str}_模型A.txt")
            self.result_A["代码"].to_csv(outA, index=False, header=False, encoding="utf-8")
        if self.result_B is not None and not self.result_B.empty:
            outB = Path(f"{date_str}_模型B.txt")
            self.result_B["代码"].to_csv(outB, index=False, header=False, encoding="utf-8")

        if (self.result_A is None or self.result_A.empty) and (self.result_B is None or self.result_B.empty):
            messagebox.showwarning("提示", "没有结果可导出。")
            return

        messagebox.showinfo("导出完成", f"同花顺格式 TXT 文件已生成在当前目录。")

if __name__ == "__main__":
    root = tk.Tk()
    app = StockAssistantApp(root)
    root.mainloop()
