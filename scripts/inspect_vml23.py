"""Deep-inspect VML 2023 raw to design 取數 (include/exclude) + NG/V rules.

VML 23's reported numbers look wrong — the 2023 sheet carries many project-team flag/adjustment
columns (計入…未實際發生 / 期後調整 / 非承批博彩 / 不符合吸引外國客源 / 未認可新增項目 …) that decide what
actually counts, plus 4 amount columns. This dumps, for each, value × Σ(調整後金額) so we can see which
rows to drop and which amount is the real one. Also flags duplicate KPMG refs (double-count).

  python scripts/inspect_vml23.py
"""
from __future__ import annotations
import sys, glob
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent


def num(s):
    return pd.to_numeric(s.astype(str).str.replace(",", "", regex=False)
                         .str.replace("^-$", "0", regex=True), errors="coerce").fillna(0)


def main():
    cands = glob.glob(str(ROOT / "data" / "**" / "vml_2023.xlsx"), recursive=True)
    print("file candidates:", cands)
    if not cands:
        print("X vml_2023.xlsx not found under data/"); return
    df = pd.read_excel(cands[0], sheet_name="2023JE")
    print(f"rows={len(df):,}  cols={len(df.columns)}")
    print("\n=== ALL COLUMNS ===")
    print("  " + " | ".join(repr(str(c)) for c in df.columns))

    amt_col = "調整後金額" if "調整後金額" in df.columns else None
    amt = num(df[amt_col]) if amt_col else pd.Series(0.0, index=df.index)

    print("\n=== amount columns (Σ + nonzero count) ===")
    for c in ["Amount", "調整數", "調整金額", "調整後金額", "MOP Amt"]:
        if c in df.columns:
            s = num(df[c])
            print(f"  {c!r:14s} Σ={s.sum():>16,.0f}  nonzero={int((s != 0).sum()):,}")

    def dist(col, topn=12):
        if col not in df.columns:
            return
        key = df[col].astype(str).str.strip().replace("", "(blank)").replace("nan", "(blank)")
        g = pd.DataFrame({"k": key, "_a": amt}).groupby("k")["_a"].agg(["size", "sum"])
        g = g.reindex(g["sum"].abs().sort_values(ascending=False).index)
        print(f"\n  COL {col!r}  ({len(g)} distinct):")
        for v, r in g.head(topn).iterrows():
            print(f"     {str(v)[:46]:48s} n={int(r['size']):6,d}  Σ={r['sum']:>15,.0f}")

    print("\n=== project-team FLAG / ADJUSTMENT columns (value × Σ調整後金額) ===")
    FLAG_KW = ["計入", "期後", "不符合", "未認可", "未實際", "非承批", "Payment", "內部資源",
               "人工支出", "關聯方", "會展", "折扣", "新增項目", "選擇"]
    for c in df.columns:
        if any(k in str(c) for k in FLAG_KW):
            dist(c, topn=6)

    print("\n=== NG / V / H source columns (value × Σ) ===")
    for c in ["投資領域", "項目類型", "項目編號", "項目名稱", "類別1", "類別2",
              "分類1", "會計科目分類", "CAPEX/OPEX", "Infrastructure/\nProgramming"]:
        dist(c, topn=14)

    ref = "KPMG ref number"
    if ref in df.columns:
        dup = df[df[ref].astype(str).duplicated(keep=False)]
        print(f"\n=== KPMG ref dups: {len(dup):,} rows share a ref "
              f"(Σ={num(dup[amt_col]).sum():,.0f}) — potential double-count ===" if amt_col else "")


if __name__ == "__main__":
    main()
