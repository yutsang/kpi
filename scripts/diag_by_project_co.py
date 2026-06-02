"""Per-project Capex / Opex / Payroll(人工) breakdown — to check 取數 against the project team.

Reads the delivered 4_大表 (data/review/{ent}_投資方向_{yr}.xlsx) and, per project, sums:
  Capex, Opex (from final_capex_opex), 人工 (H_LABOR), and grand total.
So MGM / Galaxy by-project capex/opex/payroll can be eyeballed vs 表2 / project-team golden.

Run:
  python scripts/diag_by_project_co.py --entity mgm --year 25
  python scripts/diag_by_project_co.py --entity galaxy --year 25 --out results/galaxy_co_25.tsv
"""
from __future__ import annotations
import argparse, sys, csv
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--entity", required=True)
    p.add_argument("--year", default="25")
    p.add_argument("--out", default=None)
    p.add_argument("--topn", type=int, default=40)
    args = p.parse_args()

    f = ROOT / f"data/review/{args.entity}_投資方向_{args.year}.xlsx"
    if not f.exists():
        print(f"X {f} missing — run kedro/generate first"); return
    df = pd.read_excel(f, sheet_name="4_大表")
    df["amount_mop"] = pd.to_numeric(df.get("amount_mop"), errors="coerce").fillna(0)
    proj = df.get("project", pd.Series([""] * len(df))).astype(str)
    co = df.get("final_capex_opex", pd.Series([""] * len(df))).astype(str).str.strip()
    hid = df.get("horizontal_id", pd.Series([""] * len(df))).astype(str)

    is_capex = co.str.lower().str.startswith("capex")
    is_opex = co.str.lower().str.startswith("opex")
    is_labor = hid.eq("H_LABOR")

    g = pd.DataFrame({"project": proj, "amt": df["amount_mop"],
                      "capex": df["amount_mop"].where(is_capex, 0),
                      "opex": df["amount_mop"].where(is_opex, 0),
                      "payroll": df["amount_mop"].where(is_labor, 0)})
    agg = g.groupby("project").agg(total=("amt", "sum"), capex=("capex", "sum"),
                                   opex=("opex", "sum"), payroll=("payroll", "sum"),
                                   rows=("amt", "size")).reset_index()
    agg = agg.reindex(agg["total"].abs().sort_values(ascending=False).index)

    tot = agg[["total", "capex", "opex", "payroll"]].sum()
    print(f"\n=== [{args.entity} {args.year}] by-project Capex/Opex/人工  "
          f"(projects={len(agg):,}, Σtotal={tot['total']:,.0f}, capex={tot['capex']:,.0f}, "
          f"opex={tot['opex']:,.0f}, payroll={tot['payroll']:,.0f}) ===")
    print(f"{'project':40} {'total':>15} {'capex':>15} {'opex':>15} {'payroll':>14} {'rows':>6}")
    for _, r in agg.head(args.topn).iterrows():
        print(f"{str(r['project'])[:40]:40} {r['total']:>15,.0f} {r['capex']:>15,.0f} "
              f"{r['opex']:>15,.0f} {r['payroll']:>14,.0f} {int(r['rows']):>6}")

    out = Path(args.out) if args.out else (ROOT / "results" / f"{args.entity}_co_{args.year}.tsv")
    out.parent.mkdir(parents=True, exist_ok=True)
    agg.to_csv(out, sep="\t", index=False, encoding="utf-8-sig", quoting=csv.QUOTE_MINIMAL)
    print(f"\n→ {out}  (full per-project table; paste / compare vs 表2)")


if __name__ == "__main__":
    main()
