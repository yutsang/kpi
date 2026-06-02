"""查數 / trace — for an entity+year, show the RAW rows behind a label (V or H) or a
project/account/keyword. One command → row count + Σamount + breakdown (V×H, by project,
by account) + sample rows + a full TSV dump you can paste.

Reusable for auditing any figure, e.g. "Galaxy 私人飛機運營成本 數字唔啱 — 啲 jet 行係咪已拆數?":
  python scripts/trace_label.py --entity galaxy --year 25 --h 私人飛機運營成本
  python scripts/trace_label.py --entity galaxy --year 25 --label 飛機        # match V/H/project
  python scripts/trace_label.py --entity galaxy --year 25 --project "business jet"
  python scripts/trace_label.py --entity wynn  --year 25 --v 娛樂表演 --account 538

Filters (AND-combined; all optional except --entity):
  --h <id|label substr>   horizontal_id or horizontal_label
  --v <id|label substr>   vertical_id or vertical_label
  --label <substr>        matches V OR H (id or label) — use when you don't know which axis
  --project <substr>      project / subproject
  --account <prefix>      account_code startswith
  --desc <substr>         description / account_desc
Reads data/{ent}/output/{com}_kpi_report.parquet (run kedro first).
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parent.parent
ENTITIES = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
            "vml": "company_4", "melco": "company_5", "mgm": "company_6"}


def _col(df, *cands):
    for c in cands:
        if c and c in df.columns:
            return c
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", required=True, choices=list(ENTITIES))
    ap.add_argument("--year", default="25")
    ap.add_argument("--h", default=None)
    ap.add_argument("--v", default=None)
    ap.add_argument("--label", default=None)
    ap.add_argument("--project", default=None)
    ap.add_argument("--account", default=None)
    ap.add_argument("--desc", default=None)
    ap.add_argument("--n", type=int, default=30, help="sample rows printed")
    args = ap.parse_args()
    com = ENTITIES[args.entity]

    pq = ROOT / f"data/{args.entity}/output/{com}_kpi_report.parquet"
    if not pq.exists():
        print(f"X {pq} missing — run kedro {args.entity} first"); return
    cfg = yaml.safe_load((ROOT / f"conf/{com}/parameters.yml").read_text(encoding="utf-8"))
    cols = cfg.get("columns", {})
    import pyarrow.parquet as _pq  # strip pandas StringDtype metadata → object (avoids __from_arrow__ crash)
    df = _pq.read_table(pq).replace_schema_metadata(None).to_pandas()

    yc = _col(df, "report_period", "report_year", "years")
    if yc:
        df = df[df[yc].astype(str).str.startswith(args.year)].copy()
    amt_c = _col(df, cols.get("amount"), "amount_mop", "amount", "MOP Amt", "Reported Amount(MOP)")
    df["_amt"] = pd.to_numeric(df[amt_c], errors="coerce").fillna(0) if amt_c else 0.0
    proj_c = _col(df, "project", cols.get("project"), "SubProject_Name", "Name of Investment Project")
    sub_c = _col(df, "subproject", "Sub project")
    ac_c = _col(df, "account_code", cols.get("account_code"))
    ad_c = _col(df, "account_desc", cols.get("account_desc"))
    dn_c = _col(df, "description", cols.get("description"))
    co_c = _col(df, "final_capex_opex", cols.get("capex_opex"))
    vl, vi = _col(df, "vertical_label"), _col(df, "vertical_id")
    hl, hi = _col(df, "horizontal_label"), _col(df, "horizontal_id")

    def _has(col, val):
        return df[col].astype(str).str.contains(str(val), case=False, na=False, regex=False) if col else pd.Series(False, index=df.index)

    mask = pd.Series(True, index=df.index)
    if args.h:
        mask &= (_has(hl, args.h) | _has(hi, args.h))
    if args.v:
        mask &= (_has(vl, args.v) | _has(vi, args.v))
    if args.label:
        mask &= (_has(vl, args.label) | _has(vi, args.label) | _has(hl, args.label) |
                 _has(hi, args.label) | _has(proj_c, args.label) | _has(sub_c, args.label))
    if args.project:
        mask &= (_has(proj_c, args.project) | _has(sub_c, args.project))
    if args.account and ac_c:
        mask &= df[ac_c].astype(str).str.strip().str.startswith(str(args.account))
    if args.desc:
        mask &= (_has(dn_c, args.desc) | _has(ad_c, args.desc))

    sub = df[mask]
    f = " ".join(f"--{k} {v}" for k, v in vars(args).items() if v and k not in ("entity", "year", "n"))
    print(f"\n=== {args.entity} {args.year} | {f or '(all)'} ===")
    print(f"matched rows: {len(sub):,} / {len(df):,}   Σamount = {sub['_amt'].sum():,.0f}")
    if sub.empty:
        return
    if co_c:
        print("  capex/opex:", sub.groupby(sub[co_c].astype(str))["_amt"].agg(["size", "sum"]).to_dict("index"))

    if vl and hl:
        vh = sub.pivot_table(index=vl, columns=hl, values="_amt", aggfunc="sum", fill_value=0)
        print(f"\n— V × H (Σamount) —\n{vh.to_string()}")

    if proj_c:
        g = sub.groupby(sub[proj_c].astype(str))["_amt"].agg(["size", "sum"]).rename(
            columns={"size": "rows", "sum": "Σamt"}).sort_values("Σamt", key=lambda s: s.abs(), ascending=False)
        print(f"\n— by project (top 20) —")
        print(g.head(20).to_string())

    if ac_c:
        keys = [c for c in (ac_c, ad_c) if c]
        g2 = sub.groupby([sub[k].astype(str) for k in keys])["_amt"].agg(["size", "sum"]).rename(
            columns={"size": "rows", "sum": "Σamt"}).sort_values("Σamt", key=lambda s: s.abs(), ascending=False)
        print(f"\n— by account (top 20) —")
        print(g2.head(20).to_string())

    keep = [c for c in (proj_c, sub_c, ac_c, ad_c, dn_c, co_c, vi, hi, "_amt") if c]
    samp = sub.reindex(sub["_amt"].abs().sort_values(ascending=False).index)[keep].head(args.n)
    print(f"\n— top {args.n} rows by |amount| —")
    print(samp.to_string(index=False))

    rep = ROOT / "results" / f"trace_{args.entity}_{args.year}.tsv"
    rep.parent.mkdir(parents=True, exist_ok=True)
    sub[keep].to_csv(rep, sep="\t", index=False, encoding="utf-8-sig")
    print(f"\n→ full {len(sub):,} rows: {rep}")


if __name__ == "__main__":
    main()
