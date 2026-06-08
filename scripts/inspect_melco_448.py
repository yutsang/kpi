"""Melco 2023 has a ~448M block (14% of the year) with BLANK account_code (ledger_account_id) AND
blank account_desc, desc_eg 'Unit 3040 / Unit 328', V=博彩娛樂場優化, currently H_LABOR. Per the
user the fix must come from the ORIGINAL source column (not a step4 override) — so this dumps, for
those blank-account rows, the distinct values of EVERY column in tagged_rows, to find which raw
column actually carries the account / description for that block so the melco conf can map it.

Run (Windows):  python scripts/inspect_melco_448.py
Output: prints + results/inspect_melco_448.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
TR = ROOT / "data" / "melco" / "interim" / "company_5_tagged_rows.parquet"
AC = "ledger_account_id"     # melco conf account_code col
AMT = "Amount - Amended"


def main():
    L = ["# inspect_melco_448 — blank-account 2023 block (find the real source column)"]
    if not TR.exists():
        L.append(f"X {TR} missing"); _w(L); return
    df = pd.read_parquet(TR)
    per = next((c for c in ("report_period", "report_year", "years") if c in df.columns), None)
    if per:
        df = df[df[per].astype(str).str.startswith("23")].copy()
    amt = AMT if AMT in df.columns else next((c for c in df.columns if "Amount" in str(c)), None)
    a = pd.to_numeric(df[amt], errors="coerce").fillna(0.0) if amt else pd.Series(0.0, index=df.index)
    ac = df[AC].astype("string").fillna("").str.strip() if AC in df.columns else pd.Series("", index=df.index)
    blank = ac.eq("")
    sub = df[blank]
    L.append(f"\n23 rows={len(df):,}  blank-account rows={int(blank.sum()):,}  "
             f"Σ|amt| of blank={a.abs()[blank].sum():,.0f} ({a.abs()[blank].sum()/max(a.abs().sum(),1)*100:.1f}% of yr)")
    L.append(f"amount col={amt!r}  account col={AC!r}")
    L.append(f"\n## For the blank-account rows — every column's top values (find where the real account/desc lives):")
    for c in df.columns:
        s = sub[c].astype("string").fillna("").str.strip()
        nb = s.ne("").mean() * 100
        if nb < 1:  # skip columns also empty for this block
            continue
        nun = s[s.ne("")].nunique()
        top = " | ".join(map(str, s[s.ne("")].value_counts().head(4).index))
        L.append(f"   {str(c)[:34]:34s} nb{nb:4.0f}% uniq{nun:>5}  {top[:90]}")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_melco_448.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
