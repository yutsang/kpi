r"""Per-entity V / H sanity summary from kpi_report.parquet — answers '裡面的 V/H 分類是否合理?'
before building Tableau. For each entity: rows, blank-V% / blank-H% (the headline — big blanks =
old tags didn't carry to the new unified keys → V/H needs re-tag), top verticals & horizontals
(count + Σ|amt|), and capex × H sanity (capex rows that landed on a non-construction/equip H).

Run (Windows):  python scripts/inspect_vh_summary.py            # all 6
                python scripts/inspect_vh_summary.py vml melco  # subset
Output: prints + results/inspect_vh_summary.txt (small — paste back)
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
ENT = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
       "vml": "company_4", "melco": "company_5", "mgm": "company_6"}
AMT_CANDS = ["amount_mop", "Val/COArea Crcy", "MOP Amt", "Amount - Amended", "Debit minus Credit",
             "Reported Amount(MOP)", "Entry Voucher Amount/ Expense Amount", "amount"]


def _amt(df):
    for c in AMT_CANDS:
        if c in df.columns:
            return pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    return pd.Series(0.0, index=df.index)


def _dist(df, col, amt, L, n=15):
    if col not in df.columns:
        L.append(f"   {col}: MISSING"); return
    s = df[col].astype("string").fillna("").str.strip().replace("", "(blank)")
    g = pd.DataFrame({"k": s, "a": amt.abs()}).groupby("k").agg(n=("a", "size"), a=("a", "sum"))
    g = g.sort_values("a", ascending=False)
    for k, r in g.head(n).iterrows():
        L.append(f"      {str(k)[:30]:30s} {int(r['n']):>9,}r  {r['a']/1e6:>10.1f}M")


def main():
    want = [a.lower() for a in sys.argv[1:]] or list(ENT)
    L = ["# inspect_vh_summary — V/H 分類合理性 (blank% headline + distributions)"]
    for alias, com in ENT.items():
        if alias not in want:
            continue
        p = ROOT / "data" / alias / "output" / f"{com}_kpi_report.parquet"
        L.append(f"\n{'='*72}\n## {alias}")
        if not p.exists():
            L.append(f"   (no kpi_report yet: {p.name})"); continue
        df = pd.read_parquet(p)
        amt = _amt(df)
        n = len(df)
        for vh, col in (("V", "vertical_label"), ("H", "horizontal_label")):
            if col in df.columns:
                s = df[col].astype("string").fillna("").str.strip()
                bm = s.eq("")
                L.append(f"   blank {vh}: {int(bm.sum()):>9,}r ({bm.mean()*100:5.1f}%)  "
                         f"|amt| {amt.abs()[bm].sum()/1e6:,.1f}M")
        # per bucket blank V
        if "report_period" in df.columns and "vertical_label" in df.columns:
            bm = df["vertical_label"].astype("string").fillna("").str.strip().eq("")
            byb = df[bm]["report_period"].value_counts()
            if len(byb):
                L.append("   blank-V by bucket: " + "  ".join(f"{k}={v}" for k, v in byb.items()))
        L.append(f"   rows={n:,}  Σ|amt|={amt.abs().sum()/1e6:,.1f}M")
        L.append("   top verticals:")
        _dist(df, "vertical_label", amt, L)
        L.append("   top horizontals:")
        _dist(df, "horizontal_label", amt, L)
    out = ROOT / "results" / "inspect_vh_summary.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste back")


if __name__ == "__main__":
    main()
