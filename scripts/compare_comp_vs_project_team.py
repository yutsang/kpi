"""Compare our comp scope numbers vs project team's official 內部資源支出 figures.

User input: paste project team's table (5 comp H × NG breakdown) as JSON in
`project_team_comp.json`. This script computes our breakdown from kpi_report
and prints side-by-side comparison.

Example project_team_comp.json:
{
  "wynn_25": {
    "non_gaming_total": {
      "venue": 133.08, "fnb": 65.96, "room": 15.16, "ticket": 0.04, "other": 12.30
    }
  }
}

Run:
  python scripts/compare_comp_vs_project_team.py --entity wynn --year 25
"""
from __future__ import annotations
import argparse, json, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception: pass
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
import pandas as pd
import yaml

ENTITIES = {"galaxy":"company_1","sjm":"company_2","wynn":"company_3","vml":"company_4","melco":"company_5","mgm":"company_6"}
COMP_H_MAP = {
    "venue": "H_VENUE",
    "fnb": "H_FNB",
    "room": "H_HOTEL_ROOM",
    "ticket": "H_COMP_TICKET",
    "other": "H_COMP_OTHER",
}

def main():
    p = argparse.ArgumentParser()
    p.add_argument("--entity", required=True, choices=list(ENTITIES))
    p.add_argument("--year", default="25")
    p.add_argument("--project-team-json", default="project_team_comp.json")
    args = p.parse_args()
    com = ENTITIES[args.entity]
    key = f"{args.entity}_{args.year}"

    # Load project team numbers
    pt_path = Path(args.project_team_json)
    if pt_path.exists():
        pt = json.loads(pt_path.read_text(encoding="utf-8")).get(key, {}).get("non_gaming_total", {})
    else:
        pt = {}
        print(f"⚠️  {args.project_team_json} not found, only showing our numbers")

    parquet = Path(f"data/{args.entity}/output/{com}_kpi_report.parquet")
    df = pd.read_parquet(parquet)
    cfg = yaml.safe_load(Path(f"conf/{com}/parameters.yml").read_text(encoding="utf-8"))
    amt = cfg.get("columns",{}).get("amount","")
    if amt not in df.columns:
        for c in df.columns:
            if c.strip()==amt.strip(): amt=c; break
    ycol = next((c for c in ("report_period","report_year","Yr related","years") if c in df.columns), None)
    if ycol:
        s = df[ycol].astype(str)
        df = df[s.str.startswith(args.year) | (s==f"Yr 20{args.year}") | (s==f"20{args.year}")].copy()

    # Filter NON-gaming only
    ng_col = next((c for c in ("NG11 Category","NG11 category","ng11_category","ng_scope","項目類型","博彩項目標籤") if c in df.columns), None)
    if ng_col:
        # exclude gaming
        df_ng = df[~df[ng_col].astype(str).isin(["NG0","博彩項目","gaming","Gaming"])]
    else:
        df_ng = df

    df_ng["_amt"] = pd.to_numeric(df_ng[amt], errors="coerce").fillna(0)
    total_amt = float(df_ng["_amt"].sum())

    print(f"\n{args.entity}-{args.year} NON-GAMING comp scope comparison:")
    print(f"{'item':<10} {'pt (M)':>12} {'ours (M)':>12} {'diff (M)':>12} {'diff %':>10}")
    print("-"*60)
    our_total = 0; pt_total = 0
    for short, h in COMP_H_MAP.items():
        ours_m = float(df_ng[df_ng["horizontal_id"]==h]["_amt"].sum()) / 1e6
        pt_m = pt.get(short, 0)
        diff = ours_m - pt_m
        pct = (diff/pt_m*100) if pt_m else 0
        our_total += ours_m; pt_total += pt_m
        print(f"  {short:<8} {pt_m:>11.2f} {ours_m:>11.2f} {diff:>+11.2f} {pct:>+9.1f}%")
    print(f"  {'total':<8} {pt_total:>11.2f} {our_total:>11.2f} {our_total-pt_total:>+11.2f} {(our_total-pt_total)/pt_total*100 if pt_total else 0:>+9.1f}%")
    print(f"  non_gaming_total={total_amt/1e6:.0f}M  comp%={100*our_total*1e6/total_amt if total_amt else 0:.2f}%")

if __name__ == "__main__":
    main()
