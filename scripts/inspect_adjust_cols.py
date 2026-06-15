r"""Find the per-entity ADJUSTMENT columns (調整lv1/lv2 / 調整金額 / 調整後金額) in each entity's
raw parquet AND kpi_report parquet, so we can wire them into step5 carry + 大表 + Tableau.

The carry chain (step5 _UNIFIED_EXTRA / build_master_audit UNIFIED_EXTRA / prep_tableau) keeps cols
by EXACT name (adjustment_amount / adjusted_amount / adjust_lv1 / adjust_lv2). Native 24/25 files use
Chinese names (調整金額 / 調整後金額 / 2024年度調整 …) that don't match → silently dropped. This script
shows the ACTUAL names + how many rows are populated, split by year_bucket, in BOTH:
  - raw   : data/<ent>/interim/<com>_raw.parquet   (source names, what step5 sees)
  - report: data/<ent>/output/<com>_kpi_report.parquet (what currently reaches Tableau/大表)

Run (Windows):  python scripts/inspect_adjust_cols.py
Output: prints + results/inspect_adjust_cols.txt  (paste back)
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
ENT = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
       "vml": "company_4", "melco": "company_5", "mgm": "company_6"}
# substrings that flag an adjustment-ish column (case-insensitive)
PAT = ["調整", "adjust", "lv1", "lv2", "一級", "二級", "調整後", "調整金額", "adj_", "_adj"]
# the exact English names the carry-chain looks for today
CANON = ["adjustment_amount", "adjusted_amount", "adjust_lv1", "adjust_lv2"]


def _hits(cols):
    out = []
    for c in cols:
        cl = str(c).lower()
        if any(p.lower() in cl for p in PAT):
            out.append(c)
    return out


def _scan(L, tag, p, ycol_cands):
    if not p.exists():
        L.append(f"   [{tag}] MISSING {p.name}"); return
    try:
        df = pd.read_parquet(p)
    except Exception as e:
        L.append(f"   [{tag}] read error: {e}"); return
    hits = _hits(df.columns)
    ycol = next((c for c in ycol_cands if c in df.columns), None)
    L.append(f"   [{tag}] rows={len(df):,}  adjust-ish cols={len(hits)}  (year col={ycol!r})")
    canon_present = [c for c in CANON if c in df.columns]
    L.append(f"        canon English present: {canon_present or '(NONE — carry-chain finds nothing)'}")
    for c in hits:
        s = df[c].astype("string").fillna("").str.strip()
        nb = int(s.ne("").sum())
        # numeric-sum (if it's an amount)
        num = pd.to_numeric(s.str.replace(",", "", regex=False), errors="coerce")
        ssum = num.dropna().sum()
        by = ""
        if ycol is not None and nb:
            g = df.assign(_nb=s.ne("")).groupby(df[ycol].astype(str))["_nb"].sum()
            g = g[g > 0]
            by = "  by " + ",".join(f"{k}={int(v)}" for k, v in g.items())
        L.append(f"        {str(c)[:34]:34s} nonblank={nb:>7,}  Σnum={ssum/1e6:>9.2f}M{by}")


def main():
    L = ["# inspect_adjust_cols — 調整lv1/2 / 調整金額 / 調整後金額 per entity (raw + report)"]
    for ent, com in ENT.items():
        L.append(f"\n{'='*72}\n## {ent}")
        _scan(L, "raw", ROOT / "data" / ent / "interim" / f"{com}_raw.parquet",
              ["report_period", "report_year", "Yr related", "years", "year"])
        _scan(L, "report", ROOT / "data" / ent / "output" / f"{com}_kpi_report.parquet",
              ["report_period", "year_bucket", "report_year"])
    out = ROOT / "results" / "inspect_adjust_cols.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste back")


if __name__ == "__main__":
    main()
