"""Automated reconciliation of the delivered audit files (data/review/{ent}_投資方向_{yr}.xlsx).

Reads each file's flat 4_大表 sheet (the source of the pivot, includes SJM admin-comp inject) and
emits ONE reconciliation table — per entity × year:

  rows, Σ amount_mop, per-NG (NG0..NG11) totals, V_OTHER share, H_OTHER share,
  # rows missing V or H, # negative-total NG cells.

Use it as a quick "對數" sweep after a kedro / generate run — spot dropped money (Σ vs expectation),
V_OTHER / H_OTHER spikes (classification gaps), and blanks. Writes results/reconcile_audit.tsv.

Run:
  python scripts/reconcile_audit.py
  python scripts/reconcile_audit.py --entity sjm
"""
from __future__ import annotations
import argparse, sys, csv
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
ENTITIES = ["galaxy", "sjm", "wynn", "vml", "melco", "mgm"]
NG_ORDER = [f"NG{i}" for i in range(12)]


def _one(ent, year):
    f = ROOT / f"data/review/{ent}_投資方向_{year}.xlsx"
    if not f.exists():
        return None
    try:
        df = pd.read_excel(f, sheet_name="4_大表")
    except Exception as e:
        print(f"  [{ent} {year}] cannot read 4_大表: {e}"); return None
    if "amount_mop" not in df.columns:
        return None
    amt = pd.to_numeric(df["amount_mop"], errors="coerce").fillna(0)
    tot = float(amt.sum())
    vid = df.get("vertical_id", pd.Series([""] * len(df))).astype(str)
    hid = df.get("horizontal_id", pd.Series([""] * len(df))).astype(str)
    ng = df.get("ng_code", pd.Series([""] * len(df))).astype(str)
    row = {"entity": ent, "year": year, "rows": len(df), "total_mop": round(tot, 0)}
    per_ng = amt.groupby(ng).sum()
    for n in NG_ORDER:
        row[n] = round(float(per_ng.get(n, 0.0)), 0)
    vo = float(amt[vid.eq("V_OTHER")].sum())
    ho = float(amt[hid.eq("H_OTHER")].sum())
    row["V_OTHER_mop"] = round(vo, 0)
    row["V_OTHER_%"] = round(vo / tot * 100, 1) if tot else 0
    row["H_OTHER_mop"] = round(ho, 0)
    row["H_OTHER_%"] = round(ho / tot * 100, 1) if tot else 0
    row["miss_V"] = int((vid.str.strip().isin(["", "nan", "None"])).sum())
    row["miss_H"] = int((hid.str.strip().isin(["", "nan", "None"])).sum())
    row["neg_NG_cells"] = int((per_ng < 0).sum())
    return row


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--entity", default=None, help="one entity (default: all 6)")
    args = p.parse_args()
    ents = [args.entity] if args.entity else ENTITIES

    rows = []
    for ent in ents:
        for year in ("25", "24"):
            r = _one(ent, year)
            if r:
                rows.append(r)
    if not rows:
        print("X no data/review/*_投資方向_*.xlsx found — run kedro / generate first"); return

    out = ROOT / "results"; out.mkdir(exist_ok=True)
    cols = (["entity", "year", "rows", "total_mop"] + NG_ORDER +
            ["V_OTHER_mop", "V_OTHER_%", "H_OTHER_mop", "H_OTHER_%", "miss_V", "miss_H", "neg_NG_cells"])
    with (out / "reconcile_audit.tsv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols, delimiter="\t")
        w.writeheader(); w.writerows(rows)

    print("\n=== Audit reconciliation (data/review) ===")
    print(f"{'entity':8} {'yr':3} {'rows':>7} {'total_mop':>16} {'V_OTHER%':>9} {'H_OTHER%':>9} {'miss_V':>7} {'miss_H':>7} {'neg':>4}")
    for r in rows:
        flag = "  ⚠" if (r["V_OTHER_%"] >= 20 or r["H_OTHER_%"] >= 15 or r["miss_V"] or r["miss_H"] or r["neg_NG_cells"]) else ""
        print(f"{r['entity']:8} {r['year']:3} {r['rows']:>7,} {r['total_mop']:>16,.0f} "
              f"{r['V_OTHER_%']:>8.1f}% {r['H_OTHER_%']:>8.1f}% {r['miss_V']:>7,} {r['miss_H']:>7,} {r['neg_NG_cells']:>4}{flag}")
    print(f"\n→ results/reconcile_audit.tsv  (per-NG columns included; paste back)")


if __name__ == "__main__":
    main()
