"""Flag rows where OUR vertical's implied NG contradicts a raw NG-reference column.

Use when an entity carries BOTH a vertical theme AND a project-team NG code, and they
can disagree (e.g. Galaxy: a 餐廳 row the project team tagged NG0/gaming). Driving V off
such a column bakes the NG error into V — this script surfaces those contradictions so we
can decide which side is right.

How it works: map our vertical_id → its canonical NG, parse the NG number out of the
reference column (handles "NG8" and SJM-style "B8 美食之都" — B<n> == NG<n>), and list the
projects/rows where the two NG differ. Verticals that legitimately span many NG
(V_VENUE_PERF_SPORT_MICE, V_PROPERTY_UPGRADE) are skipped (can't contradict).

Run:
  python scripts/v_vs_ng_check.py --entity galaxy --year 25                       # ng_col defaults to "NG11 Category"
  python scripts/v_vs_ng_check.py --entity sjm    --year 25 --ng_col "項目性質"
Outputs results/{ent}_v_vs_ng_{year}.tsv (project × our_V × our_NG × raw_NG × amount).
"""
from __future__ import annotations
import argparse, re, sys, csv
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import yaml

ENT = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
       "vml": "company_4", "melco": "company_5", "mgm": "company_6"}

# our vertical_id -> its canonical NG (None = spans many NG, can't contradict → skip)
V2NG = {
    "V_GAMING_VENUE": "NG0", "V_GAMING_EQUIP": "NG0",
    "V_OVERSEAS_OFFICE": "NG1", "V_REGIONAL_TEAM": "NG1", "V_INVITE_GUEST": "NG1",
    "V_INVITE_AGENCY": "NG1", "V_OVERSEAS_WEB_SEO": "NG1", "V_OVERSEAS_ROADSHOW": "NG1",
    "V_PROMO_VIDEO": "NG1", "V_REGIONAL_SALES": "NG1",
    "V_MICE": "NG2", "V_CONCERT": "NG3", "V_SPORT_EVENT": "NG4",
    "V_ART_EXHIBITION": "NG5", "V_MUSEUM": "NG5", "V_WELLNESS": "NG6",
    "V_THEME_PARK": "NG7", "V_RESTAURANT": "NG8", "V_FOOD_EVENT": "NG8",
    "V_COMMUNITY": "NG9", "V_MARITIME": "NG10", "V_OTHER": "NG11",
    "V_VENUE_PERF_SPORT_MICE": None, "V_PROPERTY_UPGRADE": None,
}
_NG = re.compile(r"\bNG(\d{1,2})\b", re.I)
_B = re.compile(r"\bB(\d{1,2})\b")  # SJM 項目性質 "B8 美食之都" → NG8


def raw_ng(val: str):
    s = str(val or "")
    m = _NG.search(s) or _B.search(s)
    return f"NG{int(m.group(1))}" if m else None


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--entity", required=True, choices=list(ENT))
    p.add_argument("--year", default="25")
    p.add_argument("--ng_col", default="NG11 Category", help="raw column holding the NG / B code")
    args = p.parse_args()
    ent, com = args.entity, ENT[args.entity]

    cfg = yaml.safe_load(Path(f"conf/{com}/parameters.yml").read_text(encoding="utf-8"))
    cols = cfg.get("columns", {}) or {}
    pq = Path(f"data/{ent}/interim/{com}_tagged_rows.parquet")
    if not pq.exists():
        print(f"X {pq} missing — run kedro first"); sys.exit(1)
    df = pd.read_parquet(pq)
    if args.ng_col not in df.columns:
        print(f"X column {args.ng_col!r} not in tagged_rows. CJK-ish columns present:")
        for c in df.columns:
            if any("一" <= ch <= "鿿" for ch in str(c)) or "NG" in str(c):
                print("   ", c)
        sys.exit(1)

    rp = next((c for c in ("report_period", "report_year") if c in df.columns), None)
    if rp:
        df = df[df[rp].astype(str).str.strip().str.startswith(args.year)].copy()
    amt_col = next((c for c in (cols.get("amount"), "amount_mop", "Reported Amount(MOP)", "amount")
                    if c and c in df.columns), None)
    proj_col = next((c for c in (cols.get("project"), "project_name", "Project", "Project Name")
                     if c and c in df.columns), None)
    amt = pd.to_numeric(df[amt_col], errors="coerce").fillna(0) if amt_col else pd.Series(0.0, index=df.index)
    vid = df.get("vertical_id", pd.Series("", index=df.index)).astype("string").fillna("")
    our_ng = vid.map(lambda v: V2NG.get(v, "?"))
    rawng = df[args.ng_col].map(raw_ng)
    proj = df[proj_col].astype("string").fillna("") if proj_col else pd.Series("", index=df.index)

    base = pd.DataFrame({"proj": proj, "vid": vid, "our_ng": our_ng, "raw_ng": rawng, "amt": amt})
    total = amt.abs().sum()
    # consider only rows where BOTH our_ng (unambiguous) and raw_ng are known
    cmp = base[base["our_ng"].notna() & base["our_ng"].ne("?") & base["raw_ng"].notna()].copy()
    mism = cmp[cmp["our_ng"] != cmp["raw_ng"]]
    cov = cmp["amt"].abs().sum()
    mis_amt = mism["amt"].abs().sum()
    print(f"[{ent}] year={args.year}  ng_col={args.ng_col!r}")
    print(f"  comparable rows: {len(cmp):,} (|amt| {cov:,.0f})  |  rows w/ no NG-ref or ambiguous-V skipped")
    print(f"  ⚠ CONTRADICTIONS: {len(mism):,} rows, |amt| {mis_amt:,.0f}  "
          f"({mis_amt/cov*100 if cov else 0:.1f}% of comparable)")

    g = mism.groupby(["proj", "vid", "our_ng", "raw_ng"]).agg(amount=("amt", "sum"), n=("amt", "size")).reset_index()
    g = g.reindex(g["amount"].abs().sort_values(ascending=False).index)
    print(f"\n  {'project':<46} {'our_V (→NG)':<28} {'rawNG':<6} {'amount':>14}")
    for _, r in g.head(40).iterrows():
        print(f"  {str(r['proj'])[:46]:<46} {str(r['vid'])[:20]+'→'+str(r['our_ng']):<28} {r['raw_ng']:<6} {r['amount']:>14,.0f}")

    out = Path("results"); out.mkdir(exist_ok=True)
    g.to_csv(out / f"{ent}_v_vs_ng_{args.year}.tsv", sep="\t", index=False, encoding="utf-8-sig")
    print(f"\n→ results/{ent}_v_vs_ng_{args.year}.tsv  ({len(g):,} mismatch project×V rows)")


if __name__ == "__main__":
    main()
