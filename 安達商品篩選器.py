#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel KPI GUI – 區間業績 + 指定商品 + 彙總(目標保費/FYP/繳費年期/APE)
+ 加入「受理日 / 核實日」的最早/最晚日期欄位
+ 商品區新增：套用商品(直接產生報表)／清除已選／用關鍵字縮短清單／重置清單
"""

from __future__ import annotations
import re, sys
from pathlib import Path
from typing import List, Optional, Dict

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ─────────────────────────────── 讀檔 & 欄位工具 ───────────────────────────────
def read_excel_auto_header(fp: Path) -> pd.DataFrame:
    """自動判斷標題列；以字串讀入，保留前導 0；欄名標準化。"""
    tmp = pd.read_excel(fp, engine="openpyxl", nrows=2, header=None, dtype=str)
    first = tmp.iloc[0]
    unnamed_ratio = sum((str(x).startswith("Unnamed") or pd.isna(x)) for x in first) / len(first)
    header = 1 if unnamed_ratio > 0.5 else 0

    df = pd.read_excel(fp, engine="openpyxl", header=header, dtype=str)
    df.columns = [str(c).strip().replace("\u200b", "") for c in df.columns]
    df = df.dropna(how="all").reset_index(drop=True)
    return df

def find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    for want in candidates:
        for c in df.columns:
            if str(c).strip() == want:
                return c
    return None

# 可能的欄位名稱（容錯）
COLS: Dict[str, List[str]] = {
    "date":      ["核實日期(入帳日)", "核實日期 (入帳日)", "核實日期", "生效日", "入帳日"],
    "apply_dt":  ["受理日"],
    "verify_dt": ["核實日期(入帳日)", "核實日期 (入帳日)", "核實日期"],
    "product":   ["商品名稱"],
    "channel":   ["通路名稱", "保經公司", "通路"],
    "branch":    ["分行名稱", "分行"],
    "is_owner":  ["負責IS", "IS", "IS負責"],
    "fyp":       ["實收保費(FYP)", "實收保費", "FYP"],
    "ape":       ["APE"],
    "target":    ["目標保費", "目標保費(元)", "目標保費-金額"],
    "pay_years": ["繳費年期", "繳費年", "年期"],
}

def colmap(df: pd.DataFrame) -> dict:
    m = {}
    for k, cands in COLS.items():
        col = find_col(df, cands)
        if col: m[k] = col
    need = ["date", "product"]   # 基本必須有日期與商品欄
    miss = [n for n in need if n not in m]
    if miss:
        raise KeyError(f"缺少必要欄位：{miss}；目前欄位={list(df.columns)}")
    return m

def to_num(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.replace(",", "", regex=False).str.replace(" ", "", regex=False)
    return pd.to_numeric(s, errors="coerce")

def parse_date_any(s: str) -> pd.Timestamp:
    """接受 YYYY-MM 或 YYYY-MM-DD；回傳 Timestamp（月份用當月1日）。"""
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
        return pd.to_datetime(s, errors="raise")
    if re.fullmatch(r"\d{4}-\d{2}", s):
        return pd.to_datetime(s + "-01", errors="raise")
    raise ValueError("日期請用 YYYY-MM 或 YYYY-MM-DD")

def month_end(ts: pd.Timestamp) -> pd.Timestamp:
    return (ts.to_period("M").asfreq("D", "end")).to_timestamp()

# ─────────────────────────────── 核心計算 ───────────────────────────────
def build_reports(df: pd.DataFrame,
                  date_col: str,
                  start_s: str, end_s: str,
                  products_selected: List[str],
                  keyword: str) -> tuple[pd.DataFrame, dict]:
    """
    回傳 (raw_filtered, tables_dict)
    tables_dict = {
      "Summary": df1, "Product": df2, "Channel": df3, "Branch": df4, "IS": df5
    }
    """
    m = colmap(df)

    # 1) 日期篩選（使用使用者在 UI 選取的 date_col）
    dates = pd.to_datetime(df[date_col], errors="coerce")
    mask = dates.notna()
    if start_s:
        s = parse_date_any(start_s); mask &= dates >= s
    if end_s:
        e = parse_date_any(end_s)
        if re.fullmatch(r"\d{4}-\d{2}", end_s): e = month_end(e)
        mask &= dates <= e
    tmp = df.loc[mask].copy()

    # 2) 商品篩選（多選 + 關鍵字包含）
    prod_col = m["product"]
    if products_selected:
        tmp = tmp[tmp[prod_col].isin(products_selected)]
    if keyword.strip():
        kw = keyword.strip()
        tmp = tmp[tmp[prod_col].astype(str).str.contains(kw, case=False, na=False)]

    if tmp.empty:
        return tmp, {}

    # 3) 數值欄位建立（轉數字；缺欄位則用 0/NaN）
    target_col = m.get("target")
    fyp_col    = m.get("fyp")
    ape_col    = m.get("ape")
    pay_col    = m.get("pay_years")

    tmp["_TARGET"] = to_num(tmp[target_col]) if target_col else 0.0
    tmp["_FYP"]    = to_num(tmp[fyp_col])    if fyp_col    else 0.0
    tmp["_APE"]    = to_num(tmp[ape_col])    if ape_col    else 0.0
    tmp["_PAYY"]   = to_num(tmp[pay_col])    if pay_col    else pd.NA

    # 4) 受理日 / 核實日（若有就轉成 datetime，供最早/最晚統計）
    apply_col  = m.get("apply_dt")
    verify_col = m.get("verify_dt")
    tmp["_APPLY_DT"]  = pd.to_datetime(tmp[apply_col],  errors="coerce") if apply_col  else pd.NaT
    tmp["_VERIFY_DT"] = pd.to_datetime(tmp[verify_col], errors="coerce") if verify_col else pd.NaT

    # 5) 建表
    tables = {}

    # Summary（總覽）
    rows = [("件數", int(len(tmp)))]
    if target_col: rows.append(("目標保費 總額", float(tmp["_TARGET"].sum())))
    if fyp_col:    rows.append(("FYP 總額",    float(tmp["_FYP"].sum())))
    if ape_col:    rows.append(("APE 總額",    float(tmp["_APE"].sum())))
    if pay_col:
        mean_pay = pd.to_numeric(tmp["_PAYY"], errors="coerce").dropna().mean()
        rows.append(("繳費年期 平均", float(mean_pay) if pd.notna(mean_pay) else None))
    if apply_col and tmp["_APPLY_DT"].notna().any():
        rows.append(("受理日(最早)", tmp["_APPLY_DT"].min()))
        rows.append(("受理日(最晚)", tmp["_APPLY_DT"].max()))
    if verify_col and tmp["_VERIFY_DT"].notna().any():
        rows.append(("核實日(最早)", tmp["_VERIFY_DT"].min()))
        rows.append(("核實日(最晚)", tmp["_VERIFY_DT"].max()))
    tables["Summary"] = pd.DataFrame(rows, columns=["指標", "值"])

    # 共用：依鍵彙總（含最早/最晚日期）
    def agg_by(key_col: str, key_label: str):
        if key_col is None: return None
        gb = tmp.groupby(key_col, dropna=False)

        specs = {}
        if target_col: specs["目標保費"] = ("_TARGET", "sum")
        if fyp_col:    specs["FYP"]     = ("_FYP",    "sum")
        if ape_col:    specs["APE"]     = ("_APE",    "sum")
        if pay_col:    specs["繳費年期(平均)"] = ("_PAYY",   "mean")
        if apply_col and tmp["_APPLY_DT"].notna().any():
            specs["受理日(最早)"] = ("_APPLY_DT", "min")
            specs["受理日(最晚)"] = ("_APPLY_DT", "max")
        if verify_col and tmp["_VERIFY_DT"].notna().any():
            specs["核實日(最早)"] = ("_VERIFY_DT", "min")
            specs["核實日(最晚)"] = ("_VERIFY_DT", "max")

        if not specs: return None
        gdf = gb.agg(**specs)
        gdf["件數"] = gb.size()
        gdf = gdf.reset_index().rename(columns={key_col: key_label})

        for k in ["FYP", "目標保費", "APE", "件數"]:
            if k in gdf.columns:
                gdf = gdf.sort_values(k, ascending=False)
                break
        return gdf

    prod_col = m["product"]
    ch_col   = m.get("channel")
    br_col   = m.get("branch")
    is_col   = m.get("is_owner")

    tables["Product"] = agg_by(prod_col, "商品名稱")
    if ch_col: tables["Channel"] = agg_by(ch_col, "通路名稱")
    if br_col: tables["Branch"]  = agg_by(br_col, "分行名稱")
    if is_col: tables["IS"]      = agg_by(is_col,  "負責IS")

    tables = {k:v for k,v in tables.items() if v is not None}
    return tmp, tables

# ─────────────────────────────── GUI ───────────────────────────────
class KPIgui(ttk.Frame):
    def __init__(self, master: tk.Tk):
        super().__init__(master, padding=10)
        self.pack(fill="both", expand=True)
        master.title("Excel KPI GUI")
        master.geometry("1180x820")

        self.df: pd.DataFrame | None = None
        self.path: Path | None = None
        self.all_products: List[str] = []

        # 第一列：檔案
        top = ttk.Frame(self); top.pack(fill="x")
        self.file_lab = ttk.Label(top, text="尚未載入檔案"); self.file_lab.pack(side="left")
        ttk.Button(top, text="選擇檔案", command=self.open_file).pack(side="right")

        # 第二列：日期條件
        datebar = ttk.LabelFrame(self, text="日期條件"); datebar.pack(fill="x", pady=(8,4))
        ttk.Label(datebar, text="日期欄").pack(side="left", padx=(8,4))
        self.date_cb = ttk.Combobox(datebar, state="readonly", width=28); self.date_cb.pack(side="left")
        ttk.Label(datebar, text="起（YYYY-MM 或 YYYY-MM-DD）").pack(side="left", padx=(12,4))
        self.s_ent = ttk.Entry(datebar, width=14); self.s_ent.pack(side="left")
        ttk.Label(datebar, text="訖（YYYY-MM 或 YYYY-MM-DD）").pack(side="left", padx=(12,4))
        self.e_ent = ttk.Entry(datebar, width=14); self.e_ent.pack(side="left")

        # 第三列：商品條件
        prodbar = ttk.LabelFrame(self, text="商品條件（可多選；或輸入關鍵字包含）"); prodbar.pack(fill="x", pady=(6,4))
        self.prod_list = tk.Listbox(prodbar, height=10, selectmode="extended")
        self.prod_list.pack(side="left", fill="both", expand=True, padx=8, pady=6)

        rightp = ttk.Frame(prodbar); rightp.pack(side="left", fill="y", padx=8, pady=6)
        ttk.Button(rightp, text="全選", command=lambda: self._select_all(self.prod_list)).pack(fill="x", pady=2)
        ttk.Button(rightp, text="全不選", command=lambda: self._clear_sel(self.prod_list)).pack(fill="x", pady=2)

        ttk.Label(rightp, text="關鍵字（包含）：").pack(anchor="w", pady=(10,2))
        self.kw_ent = ttk.Entry(rightp, width=20); self.kw_ent.pack(fill="x")
        ttk.Button(rightp, text="用關鍵字縮短清單", command=self.filter_product_list).pack(fill="x", pady=(4,2))
        ttk.Button(rightp, text="重置清單", command=self.reset_product_list).pack(fill="x", pady=2)

        ttk.Separator(rightp, orient="horizontal").pack(fill="x", pady=10)
        ttk.Button(rightp, text="套用商品（直接產生報表）", command=self.run_report).pack(fill="x", pady=4)
        ttk.Button(rightp, text="清除已選", command=lambda: self._clear_sel(self.prod_list)).pack(fill="x", pady=2)

        # 第四列：執行 & 匯出
        action = ttk.Frame(self); action.pack(fill="x", pady=6)
        ttk.Button(action, text="產生報表", command=self.run_report).pack(side="left")
        ttk.Button(action, text="匯出 Excel", command=self.export_excel).pack(side="right")

        # 下方：分頁顯示各表
        self.tabs = ttk.Notebook(self); self.tabs.pack(fill="both", expand=True)
        self.views = {}
        for name in ["Summary", "Product", "Channel", "Branch", "IS", "Raw"]:
            frm = ttk.Frame(self.tabs); self.tabs.add(frm, text=name)
            tree = ttk.Treeview(frm, show="headings")
            tree.pack(fill="both", expand=True)
            vsb = ttk.Scrollbar(frm, orient="vertical", command=tree.yview)
            vsb.pack(side="right", fill="y"); tree.configure(yscrollcommand=vsb.set)
            self.views[name] = tree

        self.raw_filtered: pd.DataFrame | None = None
        self.tables: dict = {}

        # 支援命令列帶檔
        if len(sys.argv) > 1:
            p = Path(sys.argv[1])
            if p.exists():
                self._load_file(p)

    # ─────────── 基本動作 ───────────
    def _select_all(self, lb: tk.Listbox):
        lb.selection_clear(0, tk.END); lb.selection_set(0, tk.END)

    def _clear_sel(self, lb: tk.Listbox):
        lb.selection_clear(0, tk.END)

    def open_file(self):
        fp = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xls")])
        if not fp: return
        self._load_file(Path(fp))

    def _load_file(self, path: Path):
        try:
            self.df = read_excel_auto_header(path)
            self.path = path
            self.file_lab.config(text=path.name)

            # 日期欄（先以名稱猜）
            dcols = [c for c in self.df.columns if any(k in c for k in ["日", "期", "date", "帳"])]
            self.date_cb["values"] = dcols or list(self.df.columns)
            if dcols: self.date_cb.set(dcols[0])
            else:     self.date_cb.set(self.df.columns[0])

            # 商品清單
            prod_col = find_col(self.df, COLS["product"]) or "商品名稱"
            self.all_products = sorted(pd.Series(self.df[prod_col].astype(str).unique()).dropna().tolist())
            self.reset_product_list(select_all=True)

            messagebox.showinfo("載入成功","檔案已載入，請設定日期與商品後按『產生報表』或右側『套用商品』")
        except Exception as e:
            messagebox.showerror("讀檔失敗", str(e))

    # ─────────── 商品清單操作 ───────────
    def filter_product_list(self):
        """用關鍵字縮短左側清單（僅影響清單顯示，實際篩選仍由 run_report 根據勾選與關鍵字進行）"""
        kw = self.kw_ent.get().strip().lower()
        items = [p for p in self.all_products if kw in str(p).lower()] if kw else self.all_products
        self._fill_product_list(items, select_all=False)

    def reset_product_list(self, select_all: bool=False):
        """恢復全部商品清單"""
        self._fill_product_list(self.all_products, select_all=select_all)

    def _fill_product_list(self, items: List[str], select_all: bool=False):
        self.prod_list.delete(0, tk.END)
        for p in items:
            self.prod_list.insert(tk.END, p)
        if select_all and items:
            self._select_all(self.prod_list)

    # ─────────── 產生報表 ───────────
    def run_report(self):
        if self.df is None:
            messagebox.showwarning("尚未載入","請先載入 Excel 檔"); return
        date_col = self.date_cb.get().strip()
        s = self.s_ent.get().strip()
        e = self.e_ent.get().strip()
        sel_idx = list(self.prod_list.curselection())
        sel_prods = [self.prod_list.get(i) for i in sel_idx] if sel_idx else []
        kw = self.kw_ent.get().strip()

        try:
            raw, tables = build_reports(self.df.copy(), date_col, s, e, sel_prods, kw)
        except Exception as ex:
            messagebox.showerror("產生報表失敗", str(ex)); return

        self.raw_filtered = raw
        self.tables = tables
        for t in self.views.values(): t.delete(*t.get_children())

        if not tables:
            messagebox.showinfo("無結果","符合條件的資料為 0 筆")
            return

        # 顯示
        self._show_df("Raw", raw)
        for name, df in tables.items():
            self._show_df(name, df)

        self.tabs.select(0)
        messagebox.showinfo("完成","報表已產生，可切換分頁檢視或按右下角『匯出 Excel』")

    def _show_df(self, name: str, df: pd.DataFrame):
        tree = self.views.get(name)
        if not tree: return
        tree.delete(*tree.get_children())
        cols = list(df.columns)
        tree["columns"] = cols
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=140, anchor="w")
        for _, row in df.iterrows():
            tree.insert("", "end", values=[None if pd.isna(x) else x for x in row.tolist()])

    # ─────────── 匯出 Excel ───────────
    def export_excel(self):
        if self.raw_filtered is None or not self.tables:
            messagebox.showinfo("無可匯出資料","請先按『產生報表』或『套用商品』"); return
        out = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel","*.xlsx")],
            initialfile=(self.path.stem + "_KPI報表.xlsx") if self.path else "KPI報表.xlsx",
        )
        if not out: return
        try:
            with pd.ExcelWriter(out, engine="openpyxl") as w:
                self.tables["Summary"].to_excel(w, index=False, sheet_name="Summary")
                if "Product" in self.tables: self.tables["Product"].to_excel(w, index=False, sheet_name="Product")
                if "Channel" in self.tables: self.tables["Channel"].to_excel(w, index=False, sheet_name="Channel")
                if "Branch"  in self.tables: self.tables["Branch"].to_excel(w,  index=False, sheet_name="Branch")
                if "IS"      in self.tables: self.tables["IS"].to_excel(w,      index=False, sheet_name="IS")
                self.raw_filtered.to_excel(w, index=False, sheet_name="Raw")
            messagebox.showinfo("匯出成功", f"已儲存：{out}")
        except Exception as e:
            messagebox.showerror("匯出失敗", str(e))

# ─────────────────────────────── 入口 ───────────────────────────────
if __name__ == "__main__":
    root = tk.Tk()
    app = KPIgui(root)
    root.mainloop()
