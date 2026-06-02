"""對數 tool — compare the kedro PIPELINE's H-bucket totals against a reference, per entity+year.

Two reference modes:
  --raw <xlsx>   : a build-to-raw Excel (e.g. results/melco_H_from_raw_25.xlsx) with a
                   '項目組核對_H' column — the project-team-logic answer. (Melco/SJM/VML build-to-raw.)
  (no --raw)     : just print the pipeline H-totals (so you can eyeball vs your own pivots).

Prints a compact  H | 我們pipeline Σ | 參考 Σ | 差異  table — no giant V×H dump.

Run (Windows):
  python scripts/reconcile_pipeline_h.py --entity melco --year 25 --raw results/melco_H_from_raw_25.xlsx
  python scripts/reconcile_pipeline_h.py --entity melco --year 24 --raw results/melco_H_from_raw_24.xlsx
  python scripts/reconcile_pipeline_h.py --entity wynn  --year 25            # pipeline totals only
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
ENT = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
       "vml": "company_4", "melco": "company_5", "mgm": "company_6"}


def _read_parquet(path, cols):
    names = pq.read_schema(path).names
    keep = list(dict.fromkeys(c for c in cols if c and c in names))   # dedupe → no duplicate columns
    return pq.read_table(path, columns=keep).replace_schema_metadata(None).to_pandas(), names


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", required=True, choices=list(ENT))
    ap.add_argument("--year", default="25")
    ap.add_argument("--raw", default=None, help="build-to-raw xlsx (has 項目組核對_H); omit for pipeline-only")
    args = ap.parse_args()
    com = ENT[args.entity]
    cfg = yaml.safe_load((ROOT / f"conf/{com}/parameters.yml").read_text(encoding="utf-8"))
    amt_cfg = (cfg.get("columns", {}) or {}).get("amount")

    # ── pipeline side ──
    pqp = ROOT / f"data/{args.entity}/output/{com}_kpi_report.parquet"
    if not pqp.exists():
        print(f"X {pqp} missing — run kedro {args.entity}"); return
    AMT_CANDS = [amt_cfg, "amount_mop", "amount", "MOP Amt", "Val/COArea Crcy", "Amount - Amended",
                 "調整後金額", "Reported Amount(MOP)", "Reported Amount (MOP)", "金額", "本位幣金額",
                 "Entry Voucher Amount/ Expense Amount", "Entry Voucher Amount/Expense Amount"]
    df, names = _read_parquet(pqp, AMT_CANDS + ["horizontal_label", "report_period"])
    amt = next((c for c in AMT_CANDS if c and c in df.columns), None)
    if amt is None:
        print(f"X 搵唔到 amount 欄。kpi_report 有嘅欄（揀一個畀我 --raw 或加入 candidates）:\n  "
              + " | ".join(str(n) for n in names))
        return
    if "report_period" in df.columns:
        df = df[df["report_period"].astype(str).str.startswith(args.year)]
    df[amt] = pd.to_numeric(df[amt], errors="coerce").fillna(0)
    pipe = df.groupby(df["horizontal_label"].astype(str))[amt].sum()

    # ── reference side ──
    ref = None
    if args.raw:
        rp = ROOT / args.raw
        if not rp.exists(): rp = Path(args.raw)
        if rp.exists():
            r = pd.read_excel(rp)
            hcol = "項目組核對_H" if "項目組核對_H" in r.columns else None
            racol = next((c for c in ("Amount - Amended", "Amount Amended", amt) if c in r.columns), None)
            if hcol and racol:
                r[racol] = pd.to_numeric(r[racol], errors="coerce").fillna(0)
                r = r[r[hcol].astype(str).ne("(未取數)")]
                ref = r.groupby(r[hcol].astype(str))[racol].sum()
            else:
                print(f"! {rp.name} missing 項目組核對_H / amount col")
        else:
            print(f"! raw ref {args.raw} not found")

    # ── compact diff table ──
    print(f"\n=== {args.entity} {args.year} — H 對數 (pipeline amount='{amt}') ===")
    print(f"  pipeline rows={len(df):,}  Σ={df[amt].sum():,.0f}")
    if ref is not None:
        print(f"  reference Σ={ref.sum():,.0f}  ({args.raw})\n")
        allh = sorted(set(pipe.index) | set(ref.index), key=lambda h: -abs(pipe.get(h, 0)))
        print(f"  {'H':<14} {'pipeline':>16} {'參考(項目組)':>16} {'差異':>15}")
        tot_d = 0.0
        for h in allh:
            p, q = pipe.get(h, 0.0), ref.get(h, 0.0)
            d = p - q; tot_d += abs(d)
            flag = "  ⚠" if abs(d) > 500_000 else ("  ✓" if abs(d) < 1 else "")
            print(f"  {h:<14} {p:>16,.0f} {q:>16,.0f} {d:>15,.0f}{flag}")
        print(f"  {'Σ|差異|':<14} {'':>16} {'':>16} {tot_d:>15,.0f}")
    else:
        print()
        for h, v in pipe.sort_values(key=lambda s: s.abs(), ascending=False).items():
            print(f"  {h:<14} {v:>16,.0f}")
        print(f"  {'總計':<14} {pipe.sum():>16,.0f}")


if __name__ == "__main__":
    main()
