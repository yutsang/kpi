"""MGM 25 — locate the PM (=WD5-Patron) 'Item Type' breakdown column.

Project team wants WD5-Patron (PM, ~32M, currently all Comp其他) split by an "Item Type" col:
  Room / Hotel Front Desk / Mandarin Oriental → 酒店客房
  Food & Beverage → 餐飲
  Other / Vouchers → Comp其他   (剩下 → 其他)
But tagged_rows has NO 'Item Type' column. The values likely live in a differently-named column.
This dumps, for Source=WD5* rows, EVERY column's non-blank% + distinct + top values, plus a
targeted scan for the tell-tale Item-Type strings across all columns — so we can wire a WD5-gated
column_map on whichever column actually carries them.

Run (Windows):  python scripts/inspect_mgm_wd5_pm.py
Output: prints + results/inspect_mgm_wd5_pm.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
TR = ROOT / "data" / "mgm" / "interim" / "company_6_tagged_rows.parquet"
AMT = "Debit minus Credit"
SRC = "Source"
ITEM_KW = ["Room", "Food & Beverage", "Hotel Front Desk", "Mandarin Orient", "Vouchers", "Voucher"]


def main():
    L = ["# inspect_mgm_wd5_pm — find the WD5/PM Item-Type breakdown column"]
    if not TR.exists():
        L.append(f"X {TR} missing"); _w(L); return
    df = pd.read_parquet(TR)
    per = next((c for c in ("report_period", "report_year", "years") if c in df.columns), None)
    if per:
        df = df[df[per].astype(str).str.startswith("25")].copy()
    amt = AMT if AMT in df.columns else next((c for c in df.columns if "Debit" in str(c) or "Amount" in str(c)), None)
    a = pd.to_numeric(df[amt], errors="coerce").fillna(0.0)
    if SRC not in df.columns:
        L.append("X Source column missing"); _w(L); return
    src = df[SRC].astype("string").fillna("").str.strip()
    wd5 = src.str.startswith("WD5")
    sub = df[wd5]; asub = a.abs()[wd5]
    L.append(f"\nWD5* rows={int(wd5.sum()):,}  Σ|amt|={asub.sum():,.0f}  "
             f"Source values={src[wd5].value_counts().to_dict()}")
    cur = sub["horizontal_label"].astype("string").fillna("(blank)") if "horizontal_label" in sub.columns else None
    if cur is not None:
        L.append("  current H: " + str(asub.groupby(cur).sum().round(0).sort_values(ascending=False).to_dict()))

    # (A) every column's coverage + top values for WD5 rows
    L.append(f"\n## (A) every column on WD5 rows — nb% / uniq / top values:")
    for c in df.columns:
        s = sub[c].astype("string").fillna("").str.strip()
        nb = s.ne("").mean() * 100
        if nb < 1:
            continue
        nun = s[s.ne("")].nunique()
        top = " | ".join(f"{v}({n})" for v, n in s[s.ne("")].value_counts().head(6).items())
        L.append(f"   {str(c)[:30]:30s} nb{nb:4.0f}% uniq{nun:>5}  {top[:110]}")

    # (B) targeted: which columns contain the tell-tale Item-Type strings
    L.append(f"\n## (B) columns containing Item-Type tell-tale strings (Room/F&B/Hotel Front Desk/Mandarin/Vouchers):")
    for c in df.columns:
        s = sub[c].astype("string").fillna("")
        hits = {kw: int(s.str.contains(kw, case=False, na=False).sum()) for kw in ITEM_KW}
        hits = {k: v for k, v in hits.items() if v}
        if hits:
            L.append(f"   {str(c)[:30]:30s} {hits}")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_mgm_wd5_pm.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
