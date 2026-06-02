"""Build a reconciliation (對數) view: OUR V/H vs the project-team audit U/V.

Compares, for the 25 + 25_24SY buckets only:
  HORIZONTAL — per account_code: how the amount splits across OUR horizontal_label
               vs across THEIR 原表科目分類 (audit 4_大表 V column).
  VERTICAL   — per subproject (SP code): OUR vertical_label vs THEIR 原表類別 (U).

Keyed on genuine columns (account_code / SP code) — no fragile full-row join, and
THEIR U/V is used ONLY for 對數 reference (not as a classification input).

Inputs:
  data/{ent}/interim/{com}_tagged_rows.parquet   (our output)
  {audit}.xlsx  sheet 4_大表                       (project-team golden U/V)

Outputs (results/):
  {ent}_check_H.tsv   account_code | account_desc | total | our_H_split | their_科目_split | flag
  {ent}_check_V.tsv   sub_code | sub_name | total | our_V_split | their_類別_split | flag
  console: total reconciliation + biggest disagreements

Run (Windows):
  python scripts/build_vml_check.py --entity vml --audit vml_audit_25.xlsx
"""
from __future__ import annotations
import argparse, sys, csv
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ENT = {"galaxy":"company_1","sjm":"company_2","wynn":"company_3",
       "vml":"company_4","melco":"company_5","mgm":"company_6"}
BUCKETS_BY_YEAR = {"25": {"25", "25_24SY"}, "24": {"24", "24_23SY"}}


def find(df, *cands):
    cols = {str(c).strip(): c for c in df.columns}
    for c in cands:
        if c in cols: return cols[c]
    return None


def split_str(sub, label_col, amt_col, topn=6):
    """'<label>:<amt>k | ...' sorted by |amt|, for a sub-group."""
    g = sub.groupby(sub[label_col].astype(str).str.strip(), dropna=False)[amt_col].apply(
        lambda s: pd.to_numeric(s, errors="coerce").fillna(0).sum())
    g = g.reindex(g.abs().sort_values(ascending=False).index)
    return " | ".join(f"{k or '∅'}:{v/1000:,.0f}k" for k, v in g.head(topn).items())


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--entity", default="vml", choices=list(ENT))
    p.add_argument("--audit", required=True)
    p.add_argument("--sheet", default="4_大表")
    p.add_argument("--year", default="25", choices=["25", "24"])
    args = p.parse_args()
    ent = args.entity; com = ENT[ent]
    buckets = BUCKETS_BY_YEAR[args.year]
    out = Path("results"); out.mkdir(exist_ok=True)

    pq = Path(f"data/{ent}/interim/{com}_tagged_rows.parquet")
    if not pq.exists():
        print(f"X {pq} missing — run kedro first"); sys.exit(1)
    our = pd.read_parquet(pq)
    rp = find(our, "report_period")
    if rp:
        our = our[our[rp].astype(str).str.strip().isin(buckets)].copy()
    o_amt = find(our, "MOP Amt", "amount_mop", "original_amount")
    o_acc = find(our, "Account", "account_code")
    o_adesc = find(our, "A/C Name", "account_desc")
    o_sub = find(our, "Subproject")
    o_subn = find(our, "SubProject_Name")
    o_h = find(our, "horizontal_label"); o_v = find(our, "vertical_label")
    print(f"[{ent}] our rows (25+25_24SY) = {len(our):,}  amt_total = "
          f"{pd.to_numeric(our[o_amt], errors='coerce').sum():,.0f}")

    af = Path(args.audit)
    if not af.exists():
        for alt in (Path("results")/args.audit, Path.cwd().parent/args.audit):
            if alt.exists(): af = alt; break
    th = pd.read_excel(af, sheet_name=args.sheet)
    th.columns = [str(c).strip() for c in th.columns]
    t_amt = find(th, "amount_mop", "C")
    t_acc = find(th, "account_code")
    t_sub = find(th, "subcode", "subproject")
    t_u = find(th, "原表類別"); t_v = find(th, "原表科目分類")
    t_bucket = find(th, "year_bucket")
    if t_bucket:
        th = th[th[t_bucket].astype(str).str.strip().isin(buckets)].copy()
    print(f"      their 4_大表 rows (25+25_24SY) = {len(th):,}  amt_total = "
          f"{pd.to_numeric(th[t_amt], errors='coerce').sum():,.0f}")

    # ---- HORIZONTAL by account_code ----
    rows = []
    our_codes = our.assign(_c=our[o_acc].astype(str).str.strip())
    th_codes = th.assign(_c=th[t_acc].astype(str).str.strip())
    for code in sorted(set(our_codes["_c"]) | set(th_codes["_c"])):
        osub = our_codes[our_codes["_c"] == code]
        tsub = th_codes[th_codes["_c"] == code]
        o_total = pd.to_numeric(osub[o_amt], errors="coerce").sum() if len(osub) else 0
        adesc = str(osub[o_adesc].iloc[0])[:30] if (len(osub) and o_adesc) else ""
        our_split = split_str(osub, o_h, o_amt) if len(osub) else ""
        their_split = split_str(tsub, t_v, t_amt) if len(tsub) else ""
        # flag: our top-H concept count != their, rough divergence cue
        flag = "⚠" if (len(osub) and len(tsub) and
                       osub[o_h].nunique() != tsub[t_v].nunique()) else ""
        rows.append([code, adesc, round(o_total, 0), our_split, their_split, flag])
    rows.sort(key=lambda r: -abs(r[2]))
    with (out/f"{ent}_check_H_{args.year}.tsv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter="\t")
        w.writerow(["account_code", "account_desc", "amount", "our_H_split", "their_科目_split", "flag"])
        w.writerows(rows)

    # ---- VERTICAL by subproject (SP code) ----
    vrows = []
    our_sp = our.assign(_s=our[o_sub].astype(str).str.strip()) if o_sub else None
    th_sp = th.assign(_s=th[t_sub].astype(str).str.strip()) if t_sub else None
    if our_sp is not None and th_sp is not None:
        for sp in sorted(set(our_sp["_s"]) | set(th_sp["_s"])):
            osub = our_sp[our_sp["_s"] == sp]; tsub = th_sp[th_sp["_s"] == sp]
            o_total = pd.to_numeric(osub[o_amt], errors="coerce").sum() if len(osub) else 0
            name = str(osub[o_subn].iloc[0])[:34] if (len(osub) and o_subn) else ""
            our_split = split_str(osub, o_v, o_amt) if len(osub) else ""
            their_split = split_str(tsub, t_u, t_amt) if len(tsub) else ""
            vrows.append([sp, name, round(o_total, 0), our_split, their_split])
        vrows.sort(key=lambda r: -abs(r[2]))
        with (out/f"{ent}_check_V_{args.year}.tsv").open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter="\t")
            w.writerow(["sub_code", "sub_name", "amount", "our_V_split", "their_類別_split"])
            w.writerows(vrows)

    print(f"\n=== HORIZONTAL — top 20 account_code (our H vs their 科目) ===")
    for r in rows[:20]:
        print(f"  {r[0]:<11} {str(r[1])[:24]:<24} {r[2]/1e6:>8,.1f}M  我[{r[3][:46]}] 佢[{r[4][:46]}] {r[5]}")
    print(f"\n=== VERTICAL — top 15 subproject (our V vs their 類別) ===")
    for r in vrows[:15]:
        print(f"  {r[0]:<10} {str(r[1])[:26]:<26} {r[2]/1e6:>8,.1f}M  我[{r[3][:30]}] 佢[{r[4][:30]}]")
    print(f"\n→ results/{ent}_check_H_{args.year}.tsv  +  {ent}_check_V_{args.year}.tsv")


if __name__ == "__main__":
    main()
