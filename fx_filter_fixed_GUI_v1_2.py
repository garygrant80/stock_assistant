# -*- coding: utf-8 -*-
"""
反向策略筛选 · GUI版 v1.2
- 基于 v1.1：完全复刻 fx_filter_fixed.py 的读取/清洗/筛选逻辑
- 新增：按“代码”自动去重（保留最强一条：主力净额|换手率优先）
- 扫描：*.xls, *.xlsx, *.csv, *.xls.csv, *.xlsx.csv
- 输出：fx_filter_result_YYYYMMDD.txt（表格）、ths_import_YYYYMMDD.txt（同花顺导入）
"""
import os, glob, tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import pandas as pd
import numpy as np
import datetime as dt
import re
import unicodedata

APP_VER = "v1.2"
DATE_TAG = dt.datetime.now().strftime("%Y%m%d")

# ---------- 等宽渲染 ----------
def _char_width(ch): return 2 if unicodedata.east_asian_width(ch) in ('F','W') else 1
def _wcswidth(s): return sum(_char_width(c) for c in s)
def _ljust_zh(s, width): s='' if s is None else str(s); pad=max(0,width-_wcswidth(s)); return s+' '*pad
def pretty_table(df, cols=None, pad=2):
    if df is None or df.empty: return "(无筛选结果)"
    if cols: df = df[cols]
    df = df.copy().astype(str)
    widths = [max(len(c), df[c].apply(_wcswidth).max()) + pad for c in df.columns]
    header = ''.join(_ljust_zh(c, w) for c, w in zip(df.columns, widths))
    lines = [header]
    for _, row in df.iterrows():
        lines.append(''.join(_ljust_zh(row[c], w) for c, w in zip(df.columns, widths)))
    return '\n'.join(lines)

# ---------- 原脚本函数（1:1移植） ----------
def clean_number(val):
    if pd.isna(val): return np.nan
    if isinstance(val,str):
        s = val.replace("\ufeff","").strip().replace("%","")
        if s in ["--","—",""]: return np.nan
        if "亿" in s:
            s2 = s.replace("亿","").strip()
            try: return float(s2)
            except: return np.nan
        if "万" in s:
            s2 = s.replace("万","").strip()
            try: return float(s2)/100.0
            except: return np.nan
        try: return float(s)
        except:
            m = re.findall(r"[-+]?\d*\.\d+|\d+", s)
            return float(m[0]) if m else np.nan
    try: return float(val)
    except: return np.nan

def read_any(path):
    lower = path.lower()
    ext = os.path.splitext(path)[1].lower()
    if lower.endswith(".xls.csv") or lower.endswith(".xlsx.csv") or ext == ".csv":
        for enc in ("utf-8-sig","utf-8","gbk","ansi"):
            try:
                return pd.read_csv(path, encoding=enc)
            except Exception:
                continue
        return pd.read_csv(path)
    elif ext == ".xlsx":
        return pd.read_excel(path, engine="openpyxl")
    elif ext == ".xls":
        try:
            return pd.read_excel(path, engine="xlrd")
        except Exception:
            return pd.read_excel(path)
    else:
        raise RuntimeError(f"不支持文件类型：{ext}")

def normalize_columns(df):
    cols = []
    for c in df.columns:
        if isinstance(c, str):
            c = c.replace("\ufeff","").strip()
        cols.append(c)
    df.columns = cols
    alias = {
        "现价":"股价","价格":"股价","最新价":"股价",
        "换手":"换手率",
        "自由流通":"流通盘","自由流通股":"流通盘","流通市值(亿)":"流通盘",
        "股票名称":"名称","名称 ":"名称","股票代码":"代码",
    }
    for old,new in alias.items():
        if old in df.columns and new not in df.columns:
            df.rename(columns={old:new}, inplace=True)
    if "流通盘" not in df.columns:
        for c in df.columns:
            if isinstance(c,str) and ("自由流通" in c or "流通" in c):
                df.rename(columns={c:"流通盘"}, inplace=True)
                break
    df.rename(columns={c:c.strip() if isinstance(c,str) else c for c in df.columns}, inplace=True)
    return df

def is_num(x): return pd.notna(x) and isinstance(x, (int, float, np.floating))
def gt_or_ignore(a, b):
    if is_num(a) and is_num(b): return a > b
    if is_num(a) and not is_num(b): return True
    return False

def run_strategy_on_file(path):
    df = read_any(path)
    df = normalize_columns(df)
    df.replace(["--","—"], np.nan, inplace=True)
    num_cols = ["股价","振幅","涨幅","换手率","主力净额","净流入","流通盘","散户数量","量比"]
    for col in num_cols:
        if col in df.columns:
            df[col] = df[col].apply(clean_number)
    needed = ["代码","名称","股价","流通盘","主力净额","净流入","振幅","涨幅","换手率","量比","散户数量"]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise RuntimeError("数据缺少必要列：" + "、".join(missing))

    def ok(r):
        code = str(r.get("代码",""))
        digits = "".join(ch for ch in code if ch.isdigit())
        if digits.startswith("688"): return False
        name = str(r.get("名称",""))
        if "ST" in name: return False
        try:
            if not (r["股价"] < 20): return False
            if not (r["流通盘"] < 3.5): return False
            if not (r["主力净额"] > 1000/10000): return False
            if not gt_or_ignore(r.get("主力净额"), r.get("净流入")): return False
            if not (r["振幅"] <= 14): return False
            if not (-9 <= r["涨幅"] <= 2): return False
            if not (9 <= r["换手率"] <= 22): return False
            if not (r["量比"] < 2.8): return False
            if not (r["散户数量"] <= -20): return False
            if abs(r["涨幅"]) >= 10: return False
            return True
        except Exception:
            return False

    res = df[df.apply(ok, axis=1)].copy()
    return res

# ---------- 新增：按代码去重（保留最强一条） ----------
def dedup_by_code_keep_strongest(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "代码" not in df.columns: 
        return df
    work = df.copy()
    # 添加排序键：主力净额绝对值优先，其次换手率，最后量比反向（越小越好）
    def to_num(x):
        try: return float(x)
        except: return np.nan
    if "主力净额" in work.columns:
        work["_k1"] = work["主力净额"].abs().astype(float)
    else:
        work["_k1"] = np.nan
    if "换手率" in work.columns:
        work["_k2"] = pd.to_numeric(work["换手率"], errors="coerce")
    else:
        work["_k2"] = np.nan
    if "量比" in work.columns:
        work["_k3"] = -pd.to_numeric(work["量比"], errors="coerce")  # 越小越好 → 取负数做降序
    else:
        work["_k3"] = 0

    work = work.sort_values(["_k1","_k2","_k3"], ascending=[False,False,False])
    work = work.drop_duplicates(subset=["代码"], keep="first")
    return work.drop(columns=["_k1","_k2","_k3"], errors="ignore")

# ---------- GUI 扫描与运行 ----------
def analyze_folder(folder, out_widget):
    patterns = ["*.xls","*.xlsx","*.csv","*.xls.csv","*.xlsx.csv"]
    files = []
    for p in patterns:
        files.extend(glob.glob(os.path.join(folder, p)))
    if not files:
        messagebox.showerror("错误","未找到Excel或CSV文件"); return
    files = sorted(files, key=lambda x: os.path.getmtime(x))

    out_widget.insert("end", f"📂 发现 {len(files)} 个文件，开始筛选…\n")
    all_res = []
    ok_cols = ["代码","名称","股价","主力净额","换手率","量比","振幅","涨幅"]
    for f in files:
        try:
            res = run_strategy_on_file(f)
            if not res.empty:
                res["来源文件"] = os.path.basename(f)
                all_res.append(res)
                out_widget.insert("end", f"✅ {os.path.basename(f)} 筛出 {len(res)} 条\n")
            else:
                out_widget.insert("end", f"– {os.path.basename(f)} 无匹配\n")
        except Exception as e:
            out_widget.insert("end", f"⚠️ {os.path.basename(f)} 失败：{e}\n")

    if not all_res:
        out_widget.insert("end", "\n(无筛选结果)\n")
        messagebox.showinfo("完成","未筛出任何结果。")
        return

    result_all = pd.concat(all_res, ignore_index=True)
    result = dedup_by_code_keep_strongest(result_all)

    # 输出文件
    res_path = os.path.join(folder, f"fx_filter_result_{DATE_TAG}.txt")
    ths_path = os.path.join(folder, f"ths_import_{DATE_TAG}.txt")
    with open(res_path, "w", encoding="utf-8") as f:
        f.write(pretty_table(result[[c for c in ok_cols if c in result.columns]]))
    result["代码"].dropna().astype(str).to_csv(ths_path, index=False, header=False, encoding="utf-8-sig")

    out_widget.insert("end", f"\n✅ 分析完成（去重前 {len(result_all)} 条 → 去重后 {len(result)} 条）\n结果保存：\n{res_path}\n{ths_path}\n")
    out_widget.insert("end", pretty_table(result[[c for c in ok_cols if c in result.columns]].head(30)) + "\n")

# ---------- GUI ----------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"反向策略筛选 GUI {APP_VER}")
        self.geometry("980x640")
        self.folder_var = tk.StringVar()

        row = tk.Frame(self); row.pack(fill="x", pady=6)
        tk.Label(row, text="数据文件夹：").pack(side="left")
        tk.Entry(row, textvariable=self.folder_var, width=70).pack(side="left", padx=4)
        tk.Button(row, text="选择文件夹", command=self.pick).pack(side="left", padx=4)
        tk.Button(row, text="运行分析", command=self.run, bg="#106ebe", fg="white").pack(side="left", padx=4)
        tk.Button(row, text="清空输出", command=self.clear).pack(side="left", padx=4)

        self.txt = ScrolledText(self, bg="#1e1e1e", fg="#dcdcdc", font=("Consolas",10))
        self.txt.pack(fill="both", expand=True, padx=8, pady=8)
        self.txt.insert("end","准备就绪。\n")

    def pick(self):
        d = filedialog.askdirectory(title="选择数据文件夹")
        if d: self.folder_var.set(d)

    def run(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("错误","请选择有效文件夹"); return
        self.clear()
        analyze_folder(folder, self.txt)

    def clear(self):
        self.txt.delete("1.0","end")
        self.txt.insert("end","已清空。\n")

if __name__ == "__main__":
    App().mainloop()
