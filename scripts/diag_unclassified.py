"""List the projects whose rows are UNCLASSIFIED (empty vertical_id) or V_OTHER,
for one entity + year — so we can write targeted overrides for the big ones.

Reads data/<ent>/output/<code>_kpi_report.parquet, filters the year, keeps rows
whose vertical_id is empty/NaN or V_OTHER, groups by project, shows top-N by
amount with row count + the most common account description.

Run (Windows):
  python scripts/diag_unclassified.py --entity sjm   --year 24
  python scripts/diag_unclassified.py --entity melco --year 24 --top 40
Output: results/<ent>_<year>_unclassified.tsv  (drop into results/ for review)
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import pyarrow.parquet as pq
import yaml

ROOT = Path(__file__).resolve().parent.parent
ENTITIES = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
            "vml": "company_4", "melco": "company_5", "mgm": "company_6"}
YEAR_CANDIDATES = ("report_period", "report_year", "Yr related", "years")


def fuzzy(df, name):
    if not name: return None
    if name in df.columns: return name
    for c in df.columns:
        if str(c).strip() == str(name).strip(): return c
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", required=True, choices=sorted(ENTITIES))
    ap.add_argument("--year", required=True)            # "24" or "25"
    ap.add_argument("--top", type=int, default=30)
    args = ap.parse_args()
    ent, com, tag = args.entity, ENTITIES[args.entity], str(args.year)

    parquet = ROOT / "data" / ent / "output" / f"{com}_kpi_report.parquet"
    if not parquet.exists():
        print(f"X missing {parquet.relative_to(ROOT)}"); return
    cf = yaml.safe_load((ROOT / "conf" / com / "parameters.yml").open(encoding="utf-8"))
    cols = cf.get("columns", {}) or {}
    df = pq.read_table(parquet).replace_schema_metadata(None).to_pandas()

    ycol = next((c for c in YEAR_CANDIDATES if c in df.columns), None)
    yr = df[ycol].astype("string").fillna("")
    df = df[yr.str.startswith(tag) | (yr == f"Yr 20{tag}") | (yr == f"20{tag}")].copy()

    amt = fuzzy(df, cols.get("amount", ""))
    df["_amt"] = pd.to_numeric(df[amt], errors="coerce").fillna(0) if amt else 0.0
    v = df["vertical_id"].astype("string").fillna("").str.strip()
    unc = df[(v == "") | (v == "V_OTHER") | (v.str.lower() == "nan")]
    if unc.empty:
        print(f"[{ent} {tag}] no unclassified / V_OTHER rows — good."); return

    pcol = fuzzy(df, cols.get("project", "")) or "project"
    acol = fuzzy(df, cols.get("account_desc", ""))
    grp = unc.groupby(unc[pcol].astype("string").fillna("(blank)"))
    rows = []
    for proj, g in grp:
        top_acct = (g[acol].astype("string").fillna("").value_counts().idxmax()
                    if acol and len(g) else "")
        rows.append((proj, len(g), float(g["_amt"].sum()), top_acct))
    rows.sort(key=lambda r: -r[2])

    tot = float(unc["_amt"].sum())
    out = ROOT / "results" / f"{ent}_{tag}_unclassified.tsv"
    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("w", encoding="utf-8") as f:
        f.write(f"# {ent} {tag}: unclassified/V_OTHER = {tot:,.0f} over {len(unc):,} rows, "
                f"{len(rows)} projects\n")
        f.write("project\tn_rows\tamount\ttop_account_desc\n")
        for proj, n, a, acct in rows[:args.top]:
            f.write(f"{str(proj)[:70]}\t{n}\t{a:,.0f}\t{str(acct)[:40]}\n")
    print(f"[{ent} {tag}] {tot:,.0f} unclassified over {len(rows)} projects → {out.relative_to(ROOT)}")
    for proj, n, a, acct in rows[:12]:
        print(f"   {a:>16,.0f}  {str(proj)[:50]:50} | {str(acct)[:30]}")


if __name__ == "__main__":
    main()
