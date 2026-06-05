"""inspect_mgm_combine.py — profile the manually-merged MGM raw (mgm_25_raw.xlsx, 'combine' tab) by Source.
Shows: all columns; the Source column; per-Source row count + Σamount + which project/account/NG/分類
columns are populated (each source's shape) — so we can wire up MGM's 大表.

  python scripts/inspect_mgm_combine.py
  python scripts/inspect_mgm_combine.py --file data\mgm\raw\mgm_25_raw.xlsx --sheet combine
"""
from __future__ import annotations
import argparse, glob, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent


def numify(s):
    return pd.to_numeric(s.astype(str).str.replace(",", "", regex=False).str.replace(r"^\s*-\s*$", "0", regex=True),
                         errors="coerce")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file"); ap.add_argument("--sheet", default=None)
    a = ap.parse_args()
    f = a.file
    if not f:
        cand = glob.glob(str(ROOT / "**" / "mgm_25_raw.xlsx"), recursive=True)
        if not cand:
            print("X mgm_25_raw.xlsx not found — pass --file PATH"); return
        f = cand[0]
    print(f"file: {f}")
    xl = pd.ExcelFile(f)
    print(f"sheets: {xl.sheet_names}")
    sheet = a.sheet or next((s for s in xl.sheet_names if s.strip().lower() == "combine"), xl.sheet_names[0])
    print(f"reading sheet: {sheet!r}")
    df = xl.parse(sheet)
    df.columns = [str(c).strip() for c in df.columns]
    print(f"\nrows={len(df):,}  cols={len(df.columns)}")
    print("\n=== ALL COLUMNS ===\n  " + " | ".join(repr(c) for c in df.columns))

    amt_col, amt_sum = None, -1.0
    for c in df.columns:
        s = numify(df[c])
        if s.notna().mean() > 0.5 and s.abs().sum() > amt_sum:
            amt_col, amt_sum = c, s.abs().sum()
    df["_amt"] = numify(df[amt_col]).fillna(0) if amt_col else 0.0
    print(f"\nauto-detected amount col = {amt_col!r}  (Σ={df['_amt'].sum():,.0f})")

    src = next((c for c in df.columns if str(c).strip().lower() in ("source", "source column", "來源", "data source")), None) \
        or next((c for c in df.columns if "source" in str(c).lower() or "來源" in str(c)), None)
    print(f"Source col = {src!r}")

    def cols_like(*subs):
        return [c for c in df.columns if any(s in str(c) for s in subs)]
    print("\n=== columns by role ===")
    for role, subs in [("project/subproj", ("project", "Project", "項目", "Sub", "sub", "WBS")),
                       ("account", ("account", "Account", "科目", "GL", "Cost Element")),
                       ("NG/性質", ("NG", "性質", "投資", "領域", "範疇", "類型")),
                       ("分類/類別(H/V)", ("分類", "類別", "category", "Category")),
                       ("ref/id", ("Ref", "ref", "識別", "唯一", "ID", "Doc"))]:
        print(f"  {role:18s}: {cols_like(*subs)}")

    if src:
        print(f"\n=== per-Source profile ({df[src].nunique()} sources) ===")
        rows = df.groupby(df[src].astype(str))["_amt"].agg(n="size", amt="sum").reset_index()
        rows = rows.reindex(rows["amt"].abs().sort_values(ascending=False).index)
        for _, r in rows.iterrows():
            sv = str(r[src]); sub = df[df[src].astype(str).eq(sv)]
            filled = []
            for role, subs in [("proj", ("project", "Project", "項目", "WBS")), ("acct", ("account", "Account", "科目", "Cost Element")),
                               ("NG", ("NG", "性質", "投資", "領域")), ("分類", ("分類", "類別"))]:
                pops = [c for c in cols_like(*subs)
                        if sub[c].astype(str).str.strip().replace("nan", "").ne("").mean() > 0.3]
                if pops: filled.append(f"{role}={pops[:2]}")
            print(f"  {sv[:30]:32s} n={int(r['n']):6,d}  Σ={r['amt']:>16,.0f}  | {' '.join(filled)}")


if __name__ == "__main__":
    main()
