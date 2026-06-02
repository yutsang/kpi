"""Dump Melco project-name 'INxxx_Category' tokens → current V → so the V can be corrected.

diag_ng0_comp showed Melco entertainment/sports/culture projects (project names carry an
IN00x_Category token, e.g. 'IN003_EntertainmentShow', 'IN004_SportsEvents', 'IN005_Culture&Art')
are mis-tagged V_GAMING_* (→ NG0). The IN token is the project team's true category. This dumps
every distinct IN token with row count / amount / current vertical_label so we can build a
row_vertical_override (IN token → correct V).

Run:
  python scripts/inspect_melco_in.py --year 25
"""
from __future__ import annotations
import argparse, re, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parent.parent


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--year", default="25")
    args = p.parse_args()
    pq = ROOT / "data/melco/output/company_5_kpi_report.parquet"
    if not pq.exists():
        print(f"X {pq} missing"); return
    cfg = yaml.safe_load((ROOT / "conf/company_5/parameters.yml").read_text(encoding="utf-8"))
    proj_col = cfg["columns"].get("project", "Project name - Amended")
    amt_col = cfg["columns"].get("amount", "Amount - Amended")
    df = pd.read_parquet(pq)
    yc = next((c for c in ("report_period", "report_year", "years") if c in df.columns), None)
    if yc:
        df = df[df[yc].astype(str).str.startswith(args.year)].copy()

    src = df[proj_col].astype(str)
    # also scan project_name_cols in case the IN token lives in a code/id col
    for c in cfg["columns"].get("project_name_cols", []) or []:
        if c in df.columns:
            src = src.str.cat(df[c].astype(str), sep=" ")
    df["_in"] = src.str.extract(r"(IN\d+[_ ]?[A-Za-z&]+)", expand=False).fillna("(none)")
    df["_amt"] = pd.to_numeric(df[amt_col], errors="coerce").fillna(0)
    vlab = "vertical_label" if "vertical_label" in df.columns else None

    g = df.groupby("_in").agg(rows=("_amt", "size"), amount=("_amt", "sum")).reset_index()
    g = g.reindex(g["amount"].abs().sort_values(ascending=False).index)
    print(f"{'IN token':32}{'rows':>8}{'amount':>16}   current V (top)")
    for _, r in g.iterrows():
        sub = df[df["_in"] == r["_in"]]
        topv = sub[vlab].value_counts().head(2).index.tolist() if vlab else []
        print(f"{str(r['_in'])[:32]:32}{int(r['rows']):>8}{r['amount']:>16,.0f}   {', '.join(map(str, topv))}")
    print("\n→ paste this; I map each IN token → correct V (row_vertical_override).")


if __name__ == "__main__":
    main()
