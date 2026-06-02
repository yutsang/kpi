"""Why is Melco 24 投資方向 master wrong (NG0 not isolating gaming)?

Melco's V override keys off the IN0xx initiative code being embedded in
'Project name - Amended'. If 24's project names DON'T carry IN0xx, none of the
row_vertical_overrides fire for 24 → rows keep the step2/LLM vertical (lots wrongly
→ 博彩/NG0). This script checks that, per year bucket, from the tagged output.

It reads the step4 tagged parquet (per year bucket) and reports:
  · vertical_id distribution  (is NG0/V_GAMING over-represented in 24?)
  · how many rows have 'IN0' in the project name  (the override hook)
  · 博彩項目標籤 (gaming flag) vs vertical_id  (project-team ground truth for NG0)

Run (Windows):
  python scripts/diag_melco_ng.py
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pyarrow.parquet as pq

ROOT = Path(__file__).resolve().parent.parent
COMPANY = "company_5"
PROJ = "Project name - Amended"
GAMEFLAG = "博彩項目標籤"


def load(p):
    return pq.read_table(p).replace_schema_metadata(None).to_pandas()


def main():
    out = ROOT / "data" / COMPANY / "output"
    inter = ROOT / "data" / COMPANY / "interim"
    # find tagged parquet(s) — try common names per year bucket
    cands = sorted(list(out.glob("*tagged*.parquet")) + list(inter.glob("*tagged*.parquet"))
                   + list(out.glob("*.parquet")))
    print("parquet candidates:")
    for c in cands:
        print("  ", c.relative_to(ROOT))
    if not cands:
        print("X no parquet found under data/company_5/{output,interim}"); return

    for c in cands:
        df = load(c)
        if "vertical_id" not in df.columns:
            continue
        yb = next((col for col in ("year_bucket", "bucket", "year") if col in df.columns), None)
        print(f"\n===== {c.name}  rows={len(df):,}  year_col={yb} =====")
        # IN0xx hook coverage
        if PROJ in df.columns:
            has_in = df[PROJ].astype(str).str.contains("IN0", case=False, na=False)
            print(f"  rows whose {PROJ!r} contains 'IN0': {int(has_in.sum()):,} / {len(df):,} "
                  f"({has_in.mean()*100:.1f}%)")
            if yb:
                print("  IN0-coverage by year bucket:")
                for b, g in df.groupby(df[yb].astype(str)):
                    h = g[PROJ].astype(str).str.contains("IN0", case=False, na=False)
                    print(f"     {b:10} {int(h.sum()):>7,}/{len(g):>7,}  ({h.mean()*100:4.1f}%)")
        else:
            print(f"  ! {PROJ!r} not in columns")
        # vertical_id × year
        amt = next((col for col in ("amount_mop", "amount", "Amount - Amended") if col in df.columns), None)
        if yb and amt:
            import pandas as pd
            df["_amt"] = pd.to_numeric(df[amt], errors="coerce").fillna(0)
            piv = df.pivot_table(index="vertical_id", columns=df[yb].astype(str),
                                 values="_amt", aggfunc="sum", fill_value=0)
            print("  vertical_id × year  (Σ amount):")
            print(piv.round(0).to_string())
        # gaming flag vs vertical
        if GAMEFLAG in df.columns:
            print(f"  {GAMEFLAG} vs vertical_id (counts):")
            print(df.groupby([GAMEFLAG, "vertical_id"]).size().sort_values(ascending=False).head(20).to_string())


if __name__ == "__main__":
    main()
