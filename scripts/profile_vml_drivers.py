"""Profile the VML-25 driver columns so we can build the OUR-taxonomy rules.

All the project-team PBC helper columns already live in our tagged_rows. This
dumps the VALUE distributions + cross-tabs we need to wire the rules:

  1. 項目性質 (project nature, INPUT)  × our vertical_id   — build 項目性質→our V map
  2. 項目類型 (gaming gate)            value dist
  3. Payroll (BR) populated? on CIP (17099*) rows × our horizontal_id
                                       — confirm CIP/payroll → H_LABOR split feasibility
  4. account 60010 rows × Nature (BT)  — confirm comp-nature split feasibility
  5. the 6 account-name override targets: current horizontal_id + amount + A/C Name
  6. their own labels 類別/進一步分類/分類1/分類2 (CHECK-only reference)

NOTE: 類別/進一步分類/分類1/分類2 are the project team's FINAL categories — used
ONLY to build the reconciliation (對數) file, NOT as the classification basis.

Outputs (results/):  vml_drivers__*.tsv   + console summary

Run (Windows):
  python scripts/profile_vml_drivers.py --entity vml
"""
from __future__ import annotations
import argparse, sys, csv
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ENT = {"galaxy":"company_1","sjm":"company_2","wynn":"company_3",
       "vml":"company_4","melco":"company_5","mgm":"company_6"}
OVERRIDE_CODES = ["37375", "39848", "39150", "80180", "80390", "39923"]


def find(df, *cands, exact=True):
    cols = {str(c).strip(): c for c in df.columns}
    for cand in cands:
        if exact:
            if cand in cols: return cols[cand]
        else:
            for k, v in cols.items():
                if cand in k: return v
    return None


def dist(df, col, amt, out, name, topn=40):
    if col is None:
        print(f"  [skip {name}] column not found"); return
    g = df.assign(_v=df[col].astype(str).str.strip()).groupby("_v", dropna=False).agg(
        n=("_v", "size"), amount=(amt, lambda s: pd.to_numeric(s, errors="coerce").abs().sum())
    ).reset_index().sort_values("amount", ascending=False)
    g.to_csv(out / f"vml_drivers__{name}.tsv", sep="\t", index=False, encoding="utf-8-sig")
    print(f"\n=== {name}: {col} ({len(g)} distinct) ===")
    for _, r in g.head(topn).iterrows():
        print(f"  {str(r['_v'])[:40]:<40} n={int(r['n']):>6}  amt={r['amount']:>16,.0f}")


def crosstab(df, rowcol, colcol, amt, out, name, topn=40):
    if rowcol is None or colcol is None:
        print(f"  [skip {name}] missing column"); return
    d = df.assign(_r=df[rowcol].astype(str).str.strip(), _c=df[colcol].astype(str).str.strip())
    g = d.groupby(["_r", "_c"], dropna=False).agg(
        n=("_r", "size"), amount=(amt, lambda s: pd.to_numeric(s, errors="coerce").abs().sum())
    ).reset_index().sort_values(["_r", "amount"], ascending=[True, False])
    g.to_csv(out / f"vml_drivers__{name}.tsv", sep="\t", index=False, encoding="utf-8-sig")
    print(f"\n=== {name}: {rowcol} × {colcol} ===")
    cur = None; shown = 0
    for _, r in g.iterrows():
        if r["_r"] != cur:
            cur = r["_r"]; shown = 0
            print(f"  [{str(cur)[:34]}]")
        if shown < 4:
            print(f"      → {str(r['_c'])[:30]:<30} n={int(r['n']):>6} amt={r['amount']:>15,.0f}")
            shown += 1


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
    proj_nat = find(df, "項目性質")
    proj_typ = find(df, "項目類型")
    payroll  = find(df, "Payroll")
    nature   = find(df, "Nature")
    acct     = find(df, "Account")            # L (e.g. "17099 005", "80390 010")
    acname   = find(df, "A/C Name")
    vid      = find(df, "vertical_id")
    hid      = find(df, "horizontal_id")
    their_H1 = find(df, "類別")
    their_H2 = find(df, "進一步分類")
    their_V1 = find(df, "分類1")
    their_V2 = find(df, "分類2")
    print(f"[{ent}] {len(df):,} rows | amt={amt} proj_nat={proj_nat} payroll={payroll} nature={nature} acct={acct}")

    # 1) 項目性質 (INPUT) → build map + see current divergence
    dist(df, proj_nat, amt, out, "proj_nature")
    crosstab(df, proj_nat, vid, amt, out, "projnature_x_ourV")
    crosstab(df, proj_nat, their_V1, amt, out, "projnature_x_theirV")  # sanity (check-only)

    # 2) gaming gate
    dist(df, proj_typ, amt, out, "proj_type")

    # 3) CIP (17099*) × Payroll populated × our H
    if acct is not None:
        cip = df[df[acct].astype(str).str.strip().str.startswith("17099")].copy()
        if payroll is not None:
            _has = cip[payroll].notna() & cip[payroll].astype(str).str.strip().replace(
                {"nan": "", "None": "", "<NA>": "", "NaT": ""}).ne("")
            cip["_pay_has"] = _has.map({True: "Payroll有值", False: "空"})
            crosstab(cip, "_pay_has", hid, amt, out, "CIP_payroll_x_ourH")
            dist(cip, payroll, amt, out, "CIP_payroll_values")
        else:
            print("  [skip CIP/payroll] Payroll col not found")

    # 4) comp 60010 × Nature (BT)
    if acct is not None:
        comp = df[df[acct].astype(str).str.strip().str.startswith("60010")].copy()
        print(f"\n  comp 60010 rows: {len(comp):,}")
        dist(comp, nature, amt, out, "comp60010_nature")
        # also any other column that might carry the C&E/Hotel-Room nature
        for alt in ("Comp類型", "comp支出類型", "Comp支出"):
            c = find(comp, alt)
            if c is not None and comp[c].astype(str).str.strip().replace({"nan":""}).ne("").any():
                dist(comp, c, amt, out, f"comp60010_{alt}")

    # 5) override target codes — current H
    if acct is not None:
        rows = []
        for code in OVERRIDE_CODES:
            sub = df[df[acct].astype(str).str.strip().str.startswith(code)]
            if not len(sub): rows.append([code, "(none)", "", 0, ""]); continue
            for hv, g2 in sub.groupby(sub[hid].astype(str)):
                rows.append([code,
                             str(g2[acname].iloc[0])[:34] if acname else "",
                             hv, len(g2),
                             round(pd.to_numeric(g2[amt], errors="coerce").abs().sum(), 0)])
        with (out/"vml_drivers__override_targets.tsv").open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter="\t"); w.writerow(["code","ac_name","current_H","n","amount"])
            w.writerows(rows)
        print("\n=== override targets — current H (these are what we'll flip) ===")
        for r in rows:
            print(f"  code={r[0]:<8} {str(r[1])[:30]:<30} cur_H={str(r[2]):<16} n={r[3]:>5} amt={r[4]:>14,.0f}")

    # 5b) subproject (SP code) breakdown — to build SP→our V for ambiguous natures.
    #     Subproject is a genuine JE field; we key the vertical rules on it (NOT on
    #     the audit answer columns). One row per SP: name, nature, amount, current V.
    sub_code = find(df, "Subproject")
    sub_name = find(df, "SubProject_Name")
    if sub_code is not None:
        def _mode(s):
            vc = s.astype(str).str.strip().value_counts()
            return vc.index[0] if len(vc) else ""
        d = df.assign(_sc=df[sub_code].astype(str).str.strip())
        rows = []
        for sc, g2 in d.groupby("_sc", dropna=False):
            rows.append([sc,
                         _mode(g2[sub_name]) if sub_name else "",
                         _mode(g2[proj_nat]) if proj_nat else "",
                         _mode(g2[vid]) if vid else "",
                         len(g2),
                         round(pd.to_numeric(g2[amt], errors="coerce").abs().sum(), 0)])
        rows.sort(key=lambda r: -r[5])
        with (out/"vml_drivers__subproject_breakdown.tsv").open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter="\t")
            w.writerow(["sub_code", "sub_name", "項目性質", "current_V", "n", "amount"])
            w.writerows(rows)
        print(f"\n=== subproject breakdown ({len(rows)} SP) → vml_drivers__subproject_breakdown.tsv ===")
        for r in rows:
            print(f"  {str(r[0])[:10]:<10} {str(r[1])[:30]:<30} 性質={str(r[2])[:10]:<10} curV={str(r[3])[:22]:<22} amt={r[5]:>14,.0f}")

    # 6) their final labels (CHECK-only reference)
    dist(df, their_H1, amt, out, "theirH_類別")
    dist(df, their_H2, amt, out, "theirH_進一步分類")
    dist(df, their_V1, amt, out, "theirV_分類1")
    print("\n✓ done — results/vml_drivers__*.tsv")


if __name__ == "__main__":
    main()
