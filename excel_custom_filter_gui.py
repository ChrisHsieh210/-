#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel Custom Filter GUI – 自動偵測標題列版本
"""

from __future__ import annotations
import re, sys
from pathlib import Path
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# 參數設定 -------------------------------------------------------------
SPEC_PRODUCTS = ["安達人壽美利紅美元分紅終身壽險", "安達人壽紅運旺旺美元分紅終身壽險", "安達人壽美美得益美元利率變動型終身壽險", "安達人壽美美優美元利率變動型終身壽險", "安達人壽金多美美元利率變動型終身壽險"]
OUTPUT_COLS   = ["通路名稱", "分行名稱", "負責IS", "保單號碼", "被保人姓名", "業務員姓名","商品名稱", "年繳化保費"]
DATE_CANDID   = ["核實日期(入帳日)", "核實日期 (入帳日)", "核實日期", "入帳日"]

# 工具函式 -------------------------------------------------------------
def annualize_premium(row: pd.Series) -> float:
    factor = {"年繳":1, "半年繳":2, "季繳":4, "月繳":12}.get(str(row.get("繳別","")).strip(), 1)
    try:    prem = float(row.get("實收保費(FYP)", 0))
    except: prem = 0.0
    return prem * factor

def find_date_col(df: pd.DataFrame) -> str:
    for c in DATE_CANDID:
        for col in df.columns:
            if col.strip() == c:
                return col
    raise KeyError(f"找不到核實日期欄位，目前欄位：{df.columns.tolist()}")

def read_excel_auto_header(fp: Path) -> pd.DataFrame:
    # 先用 dtype=str 抓前兩列判斷 header
    tmp = pd.read_excel(fp, engine="openpyxl", nrows=2,
                        header=None, dtype=str)
    first_row = tmp.iloc[0]
    unnamed_ratio = sum(str(x).startswith("Unnamed") or pd.isna(x)
                        for x in first_row) / len(first_row)
    hdr = 1 if unnamed_ratio > 0.5 else 0

    # 正式讀檔：整份都以「字串」讀入，保留前導 0
    return pd.read_excel(fp, engine="openpyxl", header=hdr, dtype=str)

def custom_filter(df: pd.DataFrame, ym: str) -> pd.DataFrame:
    df.columns = [c.strip().replace("\u200b","") for c in df.columns]
    date_col = find_date_col(df)
    df["核實月份"] = pd.to_datetime(df[date_col], errors="coerce").dt.to_period("M")
    df = df[df["核實月份"] == pd.Period(ym)]
    df["年繳化保費"] = df.apply(annualize_premium, axis=1)

    cond_a = (df["商品名稱"] == "安達人壽術術平安終身健康保險") & df["保單狀態"].astype(str).str.contains("有效", na=False)
    cond_b = df["商品名稱"].isin(SPEC_PRODUCTS) & (df["年繳化保費"] >= 300_000)

    return pd.concat([df[cond_a], df[cond_b]]).drop_duplicates(subset=OUTPUT_COLS)[OUTPUT_COLS]

# GUI -----------------------------------------------------------------
class GUI(ttk.Frame):
    def __init__(self, master: tk.Tk):
        super().__init__(master, padding=10); self.pack(fill="both", expand=True)
        master.title("Excel Custom Filter GUI"); master.geometry("950x600")

        top = ttk.Frame(self); top.pack(fill="x")
        self.file_lab = ttk.Label(top, text="尚未載入檔案"); self.file_lab.pack(side="left")
        ttk.Button(top, text="選擇檔案", command=self.open_file).pack(side="right")
        ttk.Label(top, text="核實月份 YYYY-MM").pack(side="left", padx=10)
        self.ym = tk.StringVar(); ttk.Entry(top, textvariable=self.ym, width=8).pack(side="left")
        ttk.Button(top, text="執行篩選", command=self.run).pack(side="left", padx=4)

        self.tree = ttk.Treeview(self, show="headings"); self.tree.pack(fill="both", expand=True)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        vsb.pack(side="right", fill="y"); self.tree.configure(yscrollcommand=vsb.set)

        self.df: pd.DataFrame|None = None; self.path: Path|None = None

    def open_file(self):
        fp = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xls")])
        if not fp: return
        try:
            self.df = read_excel_auto_header(Path(fp))
            self.path = Path(fp); self.file_lab.config(text=self.path.name)
            messagebox.showinfo("載入成功","檔案已讀取，請輸入月份並執行篩選")
        except Exception as e:
            messagebox.showerror("讀檔失敗", str(e))

    def run(self):
        if self.df is None: messagebox.showwarning("尚未載入","請先載入檔案"); return
        ym = self.ym.get().strip()
        if not re.fullmatch(r"\d{4}-\d{2}", ym): messagebox.showwarning("格式錯誤","請輸入 YYYY-MM"); return
        try:
            res = custom_filter(self.df.copy(), ym)
        except Exception as e:
            messagebox.showerror("錯誤", str(e)); return
        if res.empty: messagebox.showinfo("無結果","符合條件 0 筆"); return

        self._show(res)
        out = self.path.with_stem(self.path.stem + "_核實篩選結果")
        res.to_excel(out, index=False, sheet_name="CustomFilter")
        messagebox.showinfo("完成", f"已匯出：{out}")

    def _show(self, df: pd.DataFrame):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        for col in df.columns:
            self.tree.heading(col, text=col); self.tree.column(col, width=150, anchor="w")
        for _, r in df.iterrows(): self.tree.insert("", "end", values=r.tolist())

# main ----------------------------------------------------------------
if __name__ == "__main__":
    root = tk.Tk(); GUI(root); root.mainloop()
