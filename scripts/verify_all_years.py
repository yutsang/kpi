"""Re-test / 對數 — one compact summary per entity × report_period from the pipeline output
parquet (data/{ent}/output/company_N_kpi_report.parquet). Dumps: total amount, row count,
未分類(blank NG)%, H_OTHER%, V_OTHER%, blank-H%, blank-V% — so deviations jump out.

NOTE: MGM 23/25 deliverable is the golden-driven xlsx (data/review/mgm_投資方向_*_golden.xlsx),
NOT this parquet. This parquet = bottom-up pipeline (24 = deliverable; 23/25 = diagnostic only).

Run (Windows):
  python scripts/verify_all_years.py
Output: prints + results/verify_all_years.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
ENTITIES = [("galaxy", "company_1"), ("sjm", "company_2"), ("wynn", "company_3"),
            ("vml", "company_4"), ("melco", "company_5"), ("mgm", "company_6")]
# per-entity raw amount col (tagged_rows keeps the original name)
AMT = {"company_1": "Reported Amount(MOP)", "company_2": "Val/COArea Crcy",
       "company_3": "Entry Voucher Amount/ Expense Amount", "company_4": "MOP Amt",
       "company_5": "Amount - Amended", "company_6": "Debit minus Credit"}


def pick(df, *names):
    for n in names:
        if n in df.columns: return n
    return None


def main():
    L = ["# verify_all_years — per entity × report_period (pipeline parquet)"]
    for alias, comp in ENTITIES:
        pq = ROOT / "data" / alias / "interim" / f"{comp}_tagged_rows.parquet"
        if not pq.exists():
            L.append(f"\n## {alias}: X {pq} missing"); continue
        df = pd.read_parquet(pq)
        amt = AMT[comp] if AMT.get(comp) in df.columns else pick(df, "amount_mop", "amount")
        per = pick(df, "report_period", "report_year", "years")
        hid = pick(df, "horizontal_id"); vid = pick(df, "vertical_id")
        ngc = pick(df, "ng_code", "ng_label", "ng11_category", "NG11 Category", "Section.1", "項目類型", "項目性質")
        if not (amt and per):
            L.append(f"\n## {alias}: X missing amount/period cols ({list(df.columns)[:12]})"); continue
        a = pd.to_numeric(df[amt], errors="coerce").fillna(0.0)
        d = pd.DataFrame({"_p": df[per].astype(str), "_a": a,
                          "_h": df[hid].astype(str) if hid else "", "_v": df[vid].astype(str) if vid else "",
                          "_ng": df[ngc].astype(str) if ngc else ""})
        L.append(f"\n## {alias}  ({len(df):,} rows, Σ={a.sum():,.0f})")
        L.append(f"   {'period':10s} {'total_MOP':>16s} {'rows':>8s}  {'未分類%':>7s} {'H_OTHER%':>8s} {'V_OTHER%':>8s} {'H空%':>6s} {'V空%':>6s}")
        for p, g in d.groupby("_p"):
            tot = g["_a"].sum(); atot = g["_a"].abs().sum() or 1
            ng_blank = g.loc[g["_ng"].isin(["", "nan", "None", "未分類"]), "_a"].abs().sum() / atot * 100
            h_other = g.loc[g["_h"].eq("H_OTHER"), "_a"].abs().sum() / atot * 100
            v_other = g.loc[g["_v"].eq("V_OTHER"), "_a"].abs().sum() / atot * 100
            h_blank = g.loc[g["_h"].isin(["", "nan", "None"]), "_a"].abs().sum() / atot * 100
            v_blank = g.loc[g["_v"].isin(["", "nan", "None"]), "_a"].abs().sum() / atot * 100
            L.append(f"   {p:10s} {tot:>16,.0f} {len(g):>8,}  {ng_blank:>6.1f}% {h_other:>7.1f}% {v_other:>7.1f}% {h_blank:>5.1f}% {v_blank:>5.1f}%")
    out = ROOT / "results" / "verify_all_years.txt"
    out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
