"""SJM 2023 'serious problem' diagnosis (user: 'all 人工成本, the two capex are 0').
Dumps the sjm 2023 tagged rows: H distribution by |amount|, the Capex/Opex flag distribution,
Capex/Opex × H crosstab, and V distribution — to see whether (a) capex is genuinely 0 (everything
flagged Opex), (b) H is mis-concentrated, or (c) the deliverable diverges from tagged_rows.

Run (Windows):  python scripts/inspect_sjm_23.py
Output: prints + results/inspect_sjm_23.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
TR = ROOT / "data" / "sjm" / "interim" / "company_2_tagged_rows.parquet"
AMT = "Val/COArea Crcy"


def main():
    L = ["# inspect_sjm_23 — capex/opex + H + V diagnosis"]
    if not TR.exists():
        L.append(f"X {TR} missing"); _w(L); return
    df = pd.read_parquet(TR)
    per = next((c for c in ("report_period", "report_year", "years") if c in df.columns), None)
    if per:
        df = df[df[per].astype(str).str.startswith("23")].copy()
    amt = AMT if AMT in df.columns else next((c for c in df.columns if "Crcy" in str(c) or "Amount" in str(c)), None)
    a = pd.to_numeric(df[amt], errors="coerce").fillna(0.0)
    tot = a.abs().sum() or 1
    L.append(f"\n23 rows={len(df):,}  Σ={a.sum():,.0f}  Σ|amt|={a.abs().sum():,.0f}  amount={amt!r}")
    L.append(f"ALL columns: {list(df.columns)}")

    def dist(col, title):
        if col not in df.columns:
            L.append(f"\n## {title}: col {col!r} MISSING"); return
        s = df[col].astype("string").fillna("(blank)")
        L.append(f"\n## {title}  ({col})  by Σ|amt|:")
        for v, x in a.abs().groupby(s).sum().sort_values(ascending=False).items():
            L.append(f"   {str(v)[:34]:34s} {x/tot*100:6.1f}%  ({x:,.0f})")

    # capex/opex flag — both the raw col and the pipeline-derived one
    for c in ("Capex/Opex", "final_capex_opex", "capex_opex"):
        if c in df.columns:
            dist(c, f"CAPEX/OPEX flag")
    dist("horizontal_label", "Horizontal (H)")
    dist("vertical_label", "Vertical (V)")
    dist("項目性質", "NG (項目性質)")

    # crosstab capex/opex × H
    cflag = next((c for c in ("final_capex_opex", "Capex/Opex", "capex_opex") if c in df.columns), None)
    if cflag and "horizontal_label" in df.columns:
        L.append(f"\n## {cflag} × Horizontal  (Σ|amt| in M):")
        ct = pd.crosstab(df[cflag].astype("string").fillna("(blank)"),
                         df["horizontal_label"].astype("string").fillna("(blank)"),
                         values=a.abs(), aggfunc="sum").fillna(0) / 1e6
        L.append(ct.round(1).to_string())
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_sjm_23.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
