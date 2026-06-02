"""Reconcile MGM pipeline 取數 vs project-team golden (results/mgm_golden_25.tsv).

Golden = per 項目序號: Payroll / CAPEX / OPEX / Total (values in THOUSANDS).
Pipeline = data/mgm/output/company_6_kpi_report.parquet, where the capex/opex split lives in the
'wd' column:  CAPEX → capex ; WD3 → payroll ; Gaming_OPEX/WD1/WD2/WD4/WD5_Patron → opex.

Matches pipeline rows to a 項目序號 via project_id (zfill 3) or a 項目<n> token in project_name,
aggregates capex/opex/payroll per 序號, joins the golden, and prints the diff so the over/under
取數 projects are obvious. Rows that map to NO 序號 are shown as (unmatched) — that bucket should
be small if the project mapping is clean.

Run:
  python scripts/reconcile_mgm_golden.py
  python scripts/reconcile_mgm_golden.py --golden results/mgm_golden_25.tsv --year 25
"""
from __future__ import annotations
import argparse, re, sys, csv
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parent.parent
OPEX_WD = {"Gaming_OPEX", "WD1", "WD2", "WD4", "WD5_Patron"}   # WD3 = payroll, CAPEX = capex


def _num(x):
    x = str(x).replace(",", "").strip()
    try:
        return float(x) * 10000.0         # golden is in 萬元 (×10,000)
    except Exception:
        return 0.0


def parse_golden(p: Path):
    out = {}
    for line in p.read_text(encoding="utf-8-sig").splitlines():
        parts = [c.strip() for c in line.split("\t")]
        if len(parts) < 6:
            continue
        sn = parts[0].strip()
        if not re.fullmatch(r"\d+", sn):
            continue
        out[sn.zfill(3)] = {"name": parts[1], "payroll": _num(parts[2]),
                            "capex": _num(parts[3]), "opex": _num(parts[4]), "total": _num(parts[5])}
    return out


def _serial_token(pid, nm):
    m = re.fullmatch(r"0*(\d+)", str(pid).strip())
    if m:
        return m.group(1).zfill(3)
    m = re.search(r"項目\s*0*(\d+)", str(nm))
    if m:
        return m.group(1).zfill(3)
    return ""


def _norm_nm(s):
    """Normalize so golden 名稱 matches pipeline project_name despite cosmetic diffs:
    strip a leading 項目NNN, unify full-width / en / em dashes to '-', drop spaces."""
    s = re.sub(r"項目\s*\d+\s*", "", str(s))
    s = re.sub(r"[－—–−]", "-", s)
    return s.replace(" ", "").replace("　", "")


def make_resolver(golden):
    """序號 from project_id / 項目NNN token; else golden 名稱 (normalized) ⊂ project_name (normalized).
    Cached per unique project name. Longest golden names first = most specific match wins."""
    name2sn = sorted(((_norm_nm(g["name"]), sn) for sn, g in golden.items() if g.get("name") and len(g["name"]) >= 4),
                     key=lambda x: -len(x[0]))
    cache = {}

    def resolve(pid, nm):
        s = _serial_token(pid, nm)
        if s and s in golden:
            return s
        key = _norm_nm(nm)
        if key in cache:
            return cache[key]
        sn = ""
        for gname, gsn in name2sn:
            if gname and gname in key:
                sn = gsn; break
        cache[key] = sn
        return sn
    return resolve


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--golden", default="results/mgm_golden_25.tsv")
    p.add_argument("--year", default="25")
    args = p.parse_args()

    gp = ROOT / args.golden
    if not gp.exists():
        print(f"X golden {gp} missing"); return
    golden = parse_golden(gp)
    print(f"golden: {len(golden)} 項目, Σtotal={sum(g['total'] for g in golden.values()):,.0f} "
          f"(capex={sum(g['capex'] for g in golden.values()):,.0f}, "
          f"opex={sum(g['opex'] for g in golden.values()):,.0f}, "
          f"payroll={sum(g['payroll'] for g in golden.values()):,.0f})")

    pq = ROOT / "data/mgm/output/company_6_kpi_report.parquet"
    if not pq.exists():
        print(f"X {pq} missing — run kedro mgm first"); return
    cfg = yaml.safe_load((ROOT / "conf/company_6/parameters.yml").read_text(encoding="utf-8"))
    cols = cfg.get("columns", {})
    df = pd.read_parquet(pq)
    wd_col = next((c for c in (cols.get("capex_opex"), "wd", "final_capex_opex") if c and c in df.columns), None)
    amt_col = cols.get("amount", "amount_mop")
    proj_col = cols.get("project", "project_name")
    if not wd_col or amt_col not in df.columns:
        print(f"X cannot find wd ({wd_col}) / amount ({amt_col}). cols={list(df.columns)[:40]}"); return

    # Where did the money go across buckets? (CIP capex routed to 24/23 would explain a 25 shortfall)
    bucket_df = None
    bc = next((c for c in ("report_period", "report_year", "years") if c in df.columns), None)
    if bc:
        amt_all = pd.to_numeric(df[amt_col], errors="coerce").fillna(0)
        wd_all = df[wd_col].astype(str).str.strip()
        by_bucket = pd.DataFrame({"b": df[bc].astype(str), "amt": amt_all,
                                  "capex": amt_all.where(wd_all.eq("CAPEX"), 0)}).groupby("b").sum()
        bucket_df = by_bucket.reset_index().rename(columns={"b": "bucket", "amt": "total"})
        print("  amount by bucket (all rows):")
        for b, r in by_bucket.iterrows():
            print(f"     {b:12} total={r['amt']:>16,.0f}  capex={r['capex']:>16,.0f}")

    yc = bc
    if yc:
        df = df[df[yc].astype(str).str.startswith(args.year)].copy()
    print(f"pipeline: wd col={wd_col!r}, year {args.year} rows={len(df):,}")

    pid = df["project_id"].astype(str) if "project_id" in df.columns else pd.Series([""] * len(df), index=df.index)
    nm = df[proj_col].astype(str)
    resolve = make_resolver(golden)
    df["_sn"] = [resolve(a, b) for a, b in zip(pid, nm)]
    wd = df[wd_col].astype(str).str.strip()
    amt = pd.to_numeric(df[amt_col], errors="coerce").fillna(0)
    df["_capex"] = amt.where(wd.eq("CAPEX"), 0)
    df["_payroll"] = amt.where(wd.eq("WD3"), 0)
    df["_opex"] = amt.where(wd.isin(OPEX_WD), 0)
    df["_total"] = amt

    g = df.groupby("_sn")[["_total", "_capex", "_opex", "_payroll"]].sum()

    rows = []
    for sn, gold in sorted(golden.items()):
        pi = g.loc[sn] if sn in g.index else pd.Series({"_total": 0, "_capex": 0, "_opex": 0, "_payroll": 0})
        rows.append([sn, gold["name"][:28], gold["total"], pi["_total"], pi["_total"] - gold["total"],
                     gold["capex"], pi["_capex"], gold["opex"], pi["_opex"], gold["payroll"], pi["_payroll"]])
    unmatched = g.loc[""] if "" in g.index else None

    cols = ["序號", "項目", "G_total", "P_total", "Δtotal", "G_capex", "P_capex",
            "G_opex", "P_opex", "G_payroll", "P_payroll"]
    rdf = pd.DataFrame([[r[0], r[1]] + [round(x) for x in r[2:]] for r in rows], columns=cols)
    # most off first; add %差 for eyeballing
    rdf["Δ%"] = (rdf["Δtotal"] / rdf["G_total"].replace(0, pd.NA) * 100).round(1)
    rdf = rdf.reindex(rdf["Δtotal"].abs().sort_values(ascending=False).index).reset_index(drop=True)

    summ = pd.DataFrame({
        "metric": ["G_total", "P_total", "Δtotal", "G_capex", "P_capex", "G_opex", "P_opex", "G_payroll", "P_payroll"],
        "amount": [rdf["G_total"].sum(), rdf["P_total"].sum(), rdf["P_total"].sum() - rdf["G_total"].sum(),
                   rdf["G_capex"].sum(), rdf["P_capex"].sum(), rdf["G_opex"].sum(), rdf["P_opex"].sum(),
                   rdf["G_payroll"].sum(), rdf["P_payroll"].sum()],
    })
    if unmatched is not None:
        summ = pd.concat([summ, pd.DataFrame({
            "metric": ["(unmatched)_total", "(unmatched)_capex", "(unmatched)_opex", "(unmatched)_payroll"],
            "amount": [round(unmatched["_total"]), round(unmatched["_capex"]),
                       round(unmatched["_opex"]), round(unmatched["_payroll"])]})], ignore_index=True)

    out = ROOT / "results" / "mgm_recon_25.xlsx"
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        summ.to_excel(w, sheet_name="0_summary", index=False)
        rdf.to_excel(w, sheet_name="1_by_項目", index=False)
        if bucket_df is not None:
            bucket_df.to_excel(w, sheet_name="2_by_bucket", index=False)

    print(f"\n{'序號':5}{'項目':30}{'G_total':>13}{'P_total':>14}{'Δtotal':>14}")
    for _, r in rdf.head(25).iterrows():
        flag = "  ⚠" if abs(r["Δtotal"]) > max(50000, abs(r["G_total"]) * 0.05) else ""
        print(f"{r['序號']:5}{str(r['項目'])[:28]:30}{r['G_total']:>13,.0f}{r['P_total']:>14,.0f}{r['Δtotal']:>14,.0f}{flag}")
    if unmatched is not None:
        print(f"\n(unmatched 項目序號) pipeline Σtotal={unmatched['_total']:,.0f}  "
              f"capex={unmatched['_capex']:,.0f} opex={unmatched['_opex']:,.0f} payroll={unmatched['_payroll']:,.0f}")
    print(f"\n→ {out}  (sheets: 0_summary / 1_by_項目 / 2_by_bucket)")


if __name__ == "__main__":
    main()
