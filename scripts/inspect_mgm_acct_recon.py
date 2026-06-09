"""MGM — reconcile ONE Ledger Account (or substring) raw-WD1 vs OUR table.

Answers: why does raw "Source=WD1 + 'Ledger Account' contains 430120" = 12.59M, but our
combined table ≈ 25M? Splits the gap by the 4 usual suspects so we see which one it is:
  (A) YEAR     — our table = 23+24+25; raw filter was 2025 only  → check the report_period rows
  (B) SOURCE   — same Ledger Account code also booked under WD2/WD4/PM…, not just WD1
  (C) SIGN     — 'Debit minus Credit' is signed (comps reverse); net 12.59M vs gross |abs| 25M
  (D) OVER-MATCH — 'contains 430120' also catches 4301200 / x430120 etc (distinct values listed)

Run (Windows):  python scripts/inspect_mgm_acct_recon.py 430120
  (no arg → 430120). REUSABLE for every project-team comment account:
  463120 / 420120 / 463140 / 463130 / 463150 / 430140 / 410140 / 412140 / 540025 / …
Output: prints + results/inspect_mgm_acct_recon.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
TR = ROOT / "data" / "mgm" / "interim" / "company_6_tagged_rows.parquet"
ACCT = "Ledger Account"
SRC = "Source"


def _amt_col(df):
    for c in ("Debit minus Credit", "amount", "amount_mop"):
        if c in df.columns:
            return c
    return None


def _stat(s):
    s = pd.to_numeric(s, errors="coerce").fillna(0.0)
    return (f"net={s.sum()/1e6:8.3f}M  |abs|={s.abs().sum()/1e6:8.3f}M  "
            f"pos={s[s > 0].sum()/1e6:8.3f}M  neg={s[s < 0].sum()/1e6:8.3f}M  n={len(s)}")


def _group(L, sub, amt, col, title, n=20):
    if col not in sub.columns:
        L.append(f"\n## by {title}: col {col!r} MISSING"); return
    keys = sub[col].astype("string").fillna("(blank)")
    L.append(f"\n## by {title} ({col}):")
    g = amt.groupby(keys)
    order = g.apply(lambda x: pd.to_numeric(x, errors="coerce").abs().sum()).sort_values(ascending=False)
    for k in order.head(n).index:
        L.append(f"   {str(k)[:24]:24s} {_stat(amt[keys == k])}")


def main():
    needle = sys.argv[1] if len(sys.argv) > 1 else "430120"
    L = [f"# inspect_mgm_acct_recon — '{ACCT}' contains '{needle}'  (raw-WD1 vs OUR table)"]
    if not TR.exists():
        L.append(f"X {TR} missing"); _w(L); return
    df = pd.read_parquet(TR)
    amtc = _amt_col(df)
    if not amtc:
        L.append("X no amount column found"); _w(L); return
    L.append(f"(amount col = {amtc!r}, table rows={len(df):,})")
    acc = df[ACCT].astype("string").fillna("") if ACCT in df.columns else pd.Series("", index=df.index)
    m = acc.str.contains(needle, regex=False, na=False)
    sub = df[m].copy()
    amt = sub[amtc]
    L.append(f"\n=== matched rows: {len(sub):,} ===")
    L.append(f"ALL Source, ALL year   {_stat(amt)}")
    if SRC in sub.columns:
        wd1 = sub[SRC].astype("string").fillna("").eq("WD1")
        L.append(f"WD1 only,   ALL year   {_stat(amt[wd1.values])}")
        if "report_period" in sub.columns:
            wd1_25 = wd1 & sub["report_period"].astype("string").fillna("").eq("25")
            L.append(f"WD1 only,   year=25    {_stat(amt[wd1_25.values])}   ← compare to raw 12.59M")

    _group(L, sub, amt, "report_period", "report_period (YEAR — suspect A)")
    _group(L, sub, amt, SRC, "Source (suspect B)")
    _group(L, sub, amt, "horizontal_label", "horizontal_label (what H we gave it)")

    L.append("\n## distinct Ledger Account values matched (suspect D — over-match):")
    vc = acc[m].value_counts()
    for v, cnt in vc.head(25).items():
        L.append(f"   {str(v)[:54]:54s} ({cnt})")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_mgm_acct_recon.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
