r"""inspect_adj_lv_raw.py — sjm/melco 24 檔調整分類欄實名（修 lv1 漏 map 用）
Run: python scripts\inspect_adj_lv_raw.py
Out: results\inspect_adj_lv_raw.txt
audit_adjust_dims 發現 sjm 24/24_23SY(⚠1,310行) + melco 24(⚠1,472行) 有調整冇 lv1 →
dump 呢啲檔全部欄名 + 含[調整/事項/類別/類型/lv/原因]欄嘅 top values。
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT  = ROOT / "results" / "inspect_adj_lv_raw.txt"
FILES = [("sjm",   "sjm_2024.xlsx",   "data"),
         ("sjm",   "sjm_2023.xlsx",   "data"),
         ("melco", "melco_2024.xlsx", "data")]
KW = ["調整", "事項", "類別", "類型", "lv", "原因", "備註", "flag", "adj"]


def main():
    L = ["# sjm/melco 24 raw 調整分類欄診斷", ""]
    for ent, fn, sheet in FILES:
        f = ROOT / "data" / ent / "raw" / fn
        L += ["", "=" * 70, f"## {fn}  sheet={sheet}"]
        if not f.exists():
            L.append(f"  !! 揾唔到 {f}"); continue
        df = pd.read_excel(f, sheet_name=sheet, dtype=str)
        df.columns = [str(c).strip() for c in df.columns]
        L.append(f"  rows={len(df):,}  cols={len(df.columns)}")
        L.append("  ALL columns:")
        for i, c in enumerate(df.columns):
            L.append(f"    [{i:>3}] {c}")
        hits = [c for c in df.columns if any(k.lower() in str(c).lower() for k in KW)]
        L.append(f"\n  調整分類相關欄 top values:")
        for c in hits:
            s = df[c].fillna("").astype(str).str.strip()
            nb = int((~s.isin(["", "nan", "None", "0", "0.0"])).sum())
            vc = s[~s.isin(["", "nan", "None", "0", "0.0"])].value_counts().head(6)
            L.append(f"    [{c}] 非空={nb:,}")
            for v, n in vc.items():
                L.append(f"        {str(v)[:56]:<58} {n:>6,}")
    OUT.parent.mkdir(exist_ok=True); OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {OUT.relative_to(ROOT)}")

if __name__ == "__main__":
    main()
