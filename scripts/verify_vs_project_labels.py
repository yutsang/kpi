"""Verify our horizontal labels against the project team's own manual labels.

The raw data usually carries the project team's manual classification (一級/二級
標籤). Their dimension differs from our 橫向 (different analysis axis), but the
RELATIONSHIP reveals whether our parameters/rules are over-fragmenting: if ONE of
their labels maps cleanly to ONE of our H → consistent; if one of their labels
scatters across many of our H (or vice versa) → either genuine cross-dimension
mismatch OR our rulebase is too "over".

Two modes:
  --list                 print all columns in tagged_rows.parquet + sample values,
                         so you identify which columns are the project-team labels
  --label-cols "A,B"     cross-tab project-team label(s) × our horizontal_label
                         (amount-weighted) → results/{ent}_label_vs_ours.tsv
                         + a concentration score per project-team label
  --year 25              optionally restrict to a report_period year bucket

Reads data/{ent}/interim/{com}_tagged_rows.parquet (raw columns + our tags).

Run (Windows):
  python scripts/verify_vs_project_labels.py --entity melco --list
  python scripts/verify_vs_project_labels.py --entity melco --label-cols "支出性質,comp_nature" --year 25
"""
from __future__ import annotations
import argparse, sys, csv
from pathlib import Path
from collections import defaultdict
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
import pandas as pd
import yaml

ENT = {"galaxy":"company_1","sjm":"company_2","wynn":"company_3",
       "vml":"company_4","melco":"company_5","mgm":"company_6"}


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--entity", required=True, choices=list(ENT))
    p.add_argument("--list", action="store_true")
    p.add_argument("--label-cols", default=None)
    p.add_argument("--year", default=None)
    args = p.parse_args()
    ent = args.entity; com = ENT[ent]

    pq = Path(f"data/{ent}/interim/{com}_tagged_rows.parquet")
    if not pq.exists():
        print(f"❌ {pq} missing (run kedro through step4 first)"); sys.exit(1)
    df = pd.read_parquet(pq)
    cfg = yaml.safe_load(Path(f"conf/{com}/parameters.yml").read_text(encoding="utf-8"))
    amt_col = cfg.get("columns",{}).get("amount","")
    if amt_col not in df.columns:
        amt_col = next((c for c in ("amount_mop","amount") if c in df.columns), None)
    df["_amt"] = pd.to_numeric(df[amt_col], errors="coerce").fillna(0) if amt_col else 0

    if args.year and "report_period" in df.columns:
        df = df[df["report_period"].astype(str).str.startswith(args.year)]

    if args.list:
        print(f"\n[{ent}] {len(df):,} rows. Columns + up to 4 sample distinct values:")
        for c in df.columns:
            try:
                vals = df[c].dropna().astype(str)
                uniq = vals[vals.str.strip()!=""].unique()[:4]
            except Exception:
                uniq = []
            n = df[c].nunique(dropna=True)
            print(f"  {c:<34} (n_uniq={n:<5}) e.g. {', '.join(str(x)[:22] for x in uniq)}")
        print("\n→ identify the 項目組 一級/二級標籤 columns, then re-run with --label-cols \"col1,col2\"")
        return

    if not args.label_cols:
        print("Specify --list or --label-cols \"A,B\""); sys.exit(1)
    cols = [c.strip() for c in args.label_cols.split(",") if c.strip()]
    our = "horizontal_label" if "horizontal_label" in df.columns else "horizontal_id"
    out = Path("results"); out.mkdir(exist_ok=True)

    for lc in cols:
        if lc not in df.columns:
            print(f"  [skip] column '{lc}' not found"); continue
        g = df.groupby([lc, our], observed=True)["_amt"].agg(lambda s: s.abs().sum()).reset_index()
        # concentration: for each project-team label, what share goes to its top our-H
        conc = []
        for lab, sub in g.groupby(lc, observed=True):
            tot = sub["_amt"].sum()
            if tot <= 0: continue
            top = sub.sort_values("_amt", ascending=False).iloc[0]
            conc.append((lab, top[our], 100*top["_amt"]/tot, sub[our].nunique(), tot))
        conc.sort(key=lambda x: -x[4])
        path = out / f"{ent}_label_vs_ours__{lc.replace('/','_')[:20]}.tsv"
        with path.open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter="\t")
            w.writerow([f"projectteam_{lc}", "top_our_H", "top_share%", "n_distinct_our_H", "amount"])
            for c in conc: w.writerow([c[0], c[1], round(c[2],1), c[3], round(c[4],0)])
        # full cross-tab too
        g.to_csv(out / f"{ent}_crosstab__{lc.replace('/','_')[:20]}.tsv", sep="\t", index=False, encoding="utf-8-sig")

        print(f"\n=== [{ent}] project-team '{lc}' × our {our} ===")
        print(f"  {'their_label':<28} {'→ our top H':<16} {'share%':>6} {'#our_H':>6} {'amount':>14}")
        for c in conc[:20]:
            flag = "  ⚠over?" if c[3] >= 4 and c[2] < 70 else ""
            print(f"  {str(c[0])[:28]:<28} {str(c[1])[:16]:<16} {c[2]:>6.0f} {c[3]:>6} {c[4]:>14,.0f}{flag}")
        print(f"  → {path}")
        print(f"\n  讀法: share% 高 + #our_H 細 = 一致（我哋無 over）；")
        print(f"        #our_H 大 + share% 低 ⚠ = 佢一個標籤散落我哋好多 H → 跨維度 OR 我哋 rule 太 over")


if __name__ == "__main__":
    main()
