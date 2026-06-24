"""inspect_raw_adj_all.py — 各 entity 2025 raw Excel 調整欄診斷
Run: python scripts\inspect_raw_adj_all.py
Out: results\inspect_raw_adj_all.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT  = ROOT / "results" / "inspect_raw_adj_all.txt"

# (alias, filename, sheet, nrows_sample)
SOURCES = [
    ("galaxy-25",  "galaxy",  "galaxy_2025.xlsx",   "Combine(clean)",         5000),
    ("sjm-25",     "sjm",     "sjm_2025.xlsx",       "data",                   5000),
    ("wynn-25",    "wynn",    "wynn_2025.xlsx",       "報告投資支出明細賬",       5000),
    ("vml-25",     "vml",     "vml_2025.xlsx",        "投資支出明細賬",           5000),
    ("melco-25",   "melco",   "melco_2025.xlsx",      "Data",                   5000),
    ("mgm-25",     "mgm",     "mgm_2025.xlsx",        "data",                   5000),
    ("mgm-24",     "mgm",     "mgm_2024.xlsx",        "data",                   5000),
    ("mgm-23",     "mgm",     "mgm_2023.xlsx",        "data",                   5000),
]

ADJ_KW = ["調整", "adjust", "Adj", "前", "後", "amt", "Amt", "AMT",
          "一級", "二級", "level", "lv1", "lv2", "Level"]


def _num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s.astype(str).str.replace(",","",regex=False), errors="coerce")


def inspect_src(alias, folder, fname, sheet, nrows, L):
    path = ROOT / "data" / folder / "raw" / fname
    L += ["", "=" * 70, f"## {alias}  →  {fname}  sheet=[{sheet}]"]
    if not path.exists():
        L.append(f"  !! 揾唔到: {path}"); return
    try:
        df = pd.read_excel(path, sheet_name=sheet, dtype=str, nrows=nrows)
        df.columns = [str(c).strip() for c in df.columns]
        L.append(f"  Sample rows: {len(df):,}  Total cols: {len(df.columns)}")
    except Exception as e:
        L.append(f"  !! Error: {e}"); return

    # All cols
    L.append(f"\n  ALL columns:")
    for i, c in enumerate(df.columns):
        L.append(f"    [{i:>3}] {c}")

    # Adjustment-related cols
    adj_cols = [c for c in df.columns if any(k.lower() in str(c).lower() for k in ADJ_KW)]
    L.append(f"\n  Adjustment-related cols ({len(adj_cols)}):")
    for c in adj_cols:
        num = _num(df[c])
        nn = int(num.notna().sum())
        s = num.sum()
        nz = int((num.fillna(0) != 0).sum())
        if nn > 0 and abs(s) > 0:
            L.append(f"    [NUMERIC] {c}  non-null={nn}  Sigma={s:,.0f}  non_zero={nz}")
        elif nn > 0:
            vc = df[c].dropna().astype(str).str.strip().value_counts(dropna=False).head(5)
            L.append(f"    [TEXT]    {c}  non-null={nn}  top: " + " | ".join(f"{v}:{n}" for v,n in vc.items()))


def main():
    L = ["# inspect_raw_adj_all — 各 entity 2025 raw Excel 調整欄診斷", ""]
    for args in SOURCES:
        inspect_src(*args, L)
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}")

if __name__ == "__main__":
    main()
