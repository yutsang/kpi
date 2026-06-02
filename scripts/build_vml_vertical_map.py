"""Build the VML-25 per-subproject vertical map (OUR taxonomy), keyed on genuine
JE columns only — NOT on the audit answer columns, NOT by reading the audit xlsx.

Baseline = the current per-subproject vertical_id (already ~95% right). We then
apply targeted corrections derived from the project-team's 表1 inputs that live in
tagged_rows (項目性質 / 項目名稱 / Subproject):

  A) 項目性質 == 博彩設施及設備的優化  → V_GAMING_EQUIP   (was sometimes V_GAMING_VENUE)
     項目性質 == 博彩娛樂場場地的優化  → V_GAMING_VENUE
  B) Subproject == SP00034 (Four Seasons L2 Retail Conversion) → V_PROPERTY_UPGRADE
  C) sub_name contains museum / 博物館 / 博物馆            → V_MUSEUM

Outputs (results/):
  vml_drivers__subproject_breakdown.tsv   — clean per-SP dump (also a Mac copy source)
  vml_vertical_SP_map.tsv                 — SP | name | 項目性質 | current_V | final_V | changed
  console: every row where final_V != current_V (the review list), by amount

Run (Windows):
  python scripts/build_vml_vertical_map.py --entity vml
"""
from __future__ import annotations
import argparse, sys, csv
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ENT = {"galaxy":"company_1","sjm":"company_2","wynn":"company_3",
       "vml":"company_4","melco":"company_5","mgm":"company_6"}


def find(df, *cands):
    cols = {str(c).strip(): c for c in df.columns}
    for c in cands:
        if c in cols: return cols[c]
    return None


def correct(sub_code, sub_name, proj_nat, current_v):
    nat = (proj_nat or "").strip()
    name = (sub_name or "")
    # A) gaming by 項目性質 (nature decides venue vs equipment)
    if nat == "博彩設施及設備的優化":
        return "V_GAMING_EQUIP"
    if nat == "博彩娛樂場場地的優化":
        return "V_GAMING_VENUE"
    # C) museum anywhere in the subproject name
    low = name.lower()
    if "museum" in low or "博物館" in name or "博物馆" in name:
        return "V_MUSEUM"
    # B) specific known misclassification
    if str(sub_code).strip().startswith("SP00034"):
        return "V_PROPERTY_UPGRADE"
    return current_v


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--entity", default="vml", choices=list(ENT))
    args = p.parse_args()
    ent = args.entity; com = ENT[ent]
    pq = Path(f"data/{ent}/interim/{com}_tagged_rows.parquet")
    if not pq.exists():
        print(f"X {pq} missing — run kedro through step4"); sys.exit(1)
    df = pd.read_parquet(pq)
    out = Path("results"); out.mkdir(exist_ok=True)

    amt = find(df, "amount_mop", "MOP Amt", "original_amount") or "MOP Amt"
    sc = find(df, "Subproject"); sn = find(df, "SubProject_Name")
    pn = find(df, "項目性質"); bz = find(df, "項目名稱")
    vid = find(df, "vertical_id")
    if sc is None or vid is None:
        print("X need Subproject + vertical_id columns"); sys.exit(1)

    def mode(s):
        vc = s.astype(str).str.strip().value_counts()
        return vc.index[0] if len(vc) else ""

    rows = []
    for code, g in df.assign(_sc=df[sc].astype(str).str.strip()).groupby("_sc", dropna=False):
        name = mode(g[sn]) if sn else ""
        nat = mode(g[pn]) if pn else ""
        bzname = mode(g[bz]) if bz else ""
        cur = mode(g[vid])
        amount = round(pd.to_numeric(g[amt], errors="coerce").abs().sum(), 0)
        final = correct(code, f"{name} {bzname}", nat, cur)
        rows.append([code, name, nat, cur, final, "CHANGED" if final != cur else "", len(g), amount])
    rows.sort(key=lambda r: -r[7])

    with (out/"vml_vertical_SP_map.tsv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter="\t")
        w.writerow(["sub_code","sub_name","項目性質","current_V","final_V","changed","n","amount"])
        w.writerows(rows)
    # also a clean breakdown copy (so there is a synced Mac copy)
    with (out/"vml_drivers__subproject_breakdown.tsv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter="\t")
        w.writerow(["sub_code","sub_name","項目性質","current_V","n","amount"])
        for r in rows: w.writerow([r[0], r[1], r[2], r[3], r[6], r[7]])

    changed = [r for r in rows if r[5]]
    chg_amt = sum(r[7] for r in changed)
    print(f"[{ent}] {len(rows)} subprojects | changed {len(changed)} | changed amount {chg_amt:,.0f}")
    print(f"  → results/vml_vertical_SP_map.tsv  +  vml_drivers__subproject_breakdown.tsv\n")
    print("=== CHANGES (final_V != current_V), by amount — review these ===")
    for r in changed:
        print(f"  {r[0]:<10} {str(r[1])[:34]:<34} 性質={str(r[2])[:8]:<8} "
              f"{r[3]:<22}→ {r[4]:<22} amt={r[7]:>14,.0f}")


if __name__ == "__main__":
    main()
