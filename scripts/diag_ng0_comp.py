"""Surface suspicious NG0 (gaming) rows that carry NON-gaming / comp horizontals.

NG0 = 博彩項目 (gaming). Comp / F&B / hotel-room / venue / ticket / performer / sponsorship /
advertising spend showing up under NG0 almost always means the row's VERTICAL is wrong (the
project was LLM-tagged V_GAMING_* but the spend is actually a non-gaming event/comp).

Reads the delivered 4_大表 sheet (data/review/{ent}_投資方向_{yr}.xlsx) and reports, per entity×year,
the NG0 amount sitting in each non-gaming H, then the worst (vertical_label, project, H) offenders —
so the V can be corrected (row_vertical_overrides or project re-tag).

Run:
  python scripts/diag_ng0_comp.py                 # all 6 entities, year 25
  python scripts/diag_ng0_comp.py --entity melco --year 25
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
ENTITIES = ["galaxy", "sjm", "wynn", "vml", "melco", "mgm"]
# Horizontals that should essentially NEVER be NG0 (gaming) — comp / event / non-gaming
NON_GAMING_H = {"H_FNB", "H_HOTEL_ROOM", "H_VENUE", "H_COMP_TICKET", "H_COMP_OTHER",
                "H_PERFORMER", "H_SPONSORSHIP", "H_ADVERTISING"}
NON_GAMING_H_LABELS = {"餐飲", "酒店客房", "活動場地", "贈票支出", "Comp其他",
                       "合約成本（演藝）", "贊助費", "廣告及推廣"}


def _one(ent, year, topn):
    f = ROOT / f"data/review/{ent}_投資方向_{year}.xlsx"
    if not f.exists():
        return
    df = pd.read_excel(f, sheet_name="4_大表")
    if "ng_code" not in df.columns or "amount_mop" not in df.columns:
        print(f"[{ent} {year}] 4_大表 missing ng_code/amount_mop"); return
    df["amount_mop"] = pd.to_numeric(df["amount_mop"], errors="coerce").fillna(0)
    ng0 = df[df["ng_code"].astype(str) == "NG0"].copy()
    if ng0.empty:
        print(f"\n[{ent} {year}] NG0 rows: 0"); return
    hid = ng0.get("horizontal_id", pd.Series([""] * len(ng0))).astype(str)
    hlab = ng0.get("horizontal_label", pd.Series([""] * len(ng0))).astype(str)
    susp = ng0[hid.isin(NON_GAMING_H) | hlab.isin(NON_GAMING_H_LABELS)].copy()
    tot_ng0 = ng0["amount_mop"].sum()
    tot_susp = susp["amount_mop"].sum()
    print(f"\n=== [{ent} {year}] NG0 Σ={tot_ng0:,.0f} | suspicious non-gaming-H Σ={tot_susp:,.0f} "
          f"({tot_susp/tot_ng0*100 if tot_ng0 else 0:.1f}%) rows={len(susp):,} ===")
    if susp.empty:
        return
    by_h = susp.groupby("horizontal_label")["amount_mop"].sum().sort_values(key=abs, ascending=False)
    print("  by H:  " + " | ".join(f"{k}={v:,.0f}" for k, v in by_h.items()))
    gcols = [c for c in ("vertical_label", "project", "horizontal_label") if c in susp.columns]
    g = (susp.groupby(gcols)["amount_mop"].agg(["sum", "size"]).reset_index()
         .rename(columns={"sum": "amount", "size": "rows"}))
    g = g.reindex(g["amount"].abs().sort_values(ascending=False).index).head(topn)
    print(f"  top {len(g)} (vertical | project | H | amount | rows):")
    for _, r in g.iterrows():
        print(f"    {str(r.get('vertical_label',''))[:14]:14} | {str(r.get('project',''))[:34]:34} | "
              f"{str(r.get('horizontal_label',''))[:10]:10} | {r['amount']:>14,.0f} | {int(r['rows'])}")


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--entity", default=None)
    p.add_argument("--year", default="25")
    p.add_argument("--topn", type=int, default=15)
    args = p.parse_args()
    for ent in ([args.entity] if args.entity else ENTITIES):
        _one(ent, args.year, args.topn)


if __name__ == "__main__":
    main()
