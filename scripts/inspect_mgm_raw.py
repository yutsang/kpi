"""Inspect the MGM raw source files (f1 capex / f2 Master / f3 gaming opex / f4 nongaming opex)
to locate the ~257M capex shortfall vs golden — runs on the machine that HAS data/mgm/raw/.

For each source file: lists every sheet (rows × cols), finds the amount column (numeric, biggest |Σ|),
prints its total. For f1 (capex) it ALSO breaks the total down by the two prebuild filter columns —
'Budget Source' and 'Ledger Hierarchy Level 4' (CIP) — so we can see if a filter is dropping capex.

Run (on Windows):
  python scripts/inspect_mgm_raw.py
  python scripts/inspect_mgm_raw.py --raw data\mgm\raw
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parent.parent


def _amt_col(df):
    """The numeric column with the largest absolute sum (the likely amount)."""
    best, bestsum = None, -1.0
    for c in df.columns:
        s = pd.to_numeric(df[c], errors="coerce")
        if s.notna().sum() < max(3, len(df) * 0.3):
            continue
        a = float(s.abs().sum())
        if a > bestsum:
            best, bestsum = c, a
    return best


def _dump_break(df, amt, col):
    if col not in df.columns:
        print(f"      ({col!r} not in sheet)"); return
    g = pd.to_numeric(df[amt], errors="coerce").fillna(0).groupby(df[col].astype(str).str.strip()).agg(["sum", "size"])
    g = g.reindex(g["sum"].abs().sort_values(ascending=False).index)
    print(f"      by {col!r}:")
    for v, r in g.head(20).iterrows():
        print(f"        {str(v)[:42]:42} {r['sum']:>16,.0f}  ({int(r['size'])})")


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--raw", default="data/mgm/raw")
    args = p.parse_args()
    raw = ROOT / args.raw
    cfg = yaml.safe_load((ROOT / "conf/company_6/parameters.yml").read_text(encoding="utf-8"))
    src = cfg.get("prebuild_sources") or {}
    if not src:
        print("X no prebuild_sources in conf"); return

    for key, fname in src.items():
        fp = raw / fname
        print(f"\n{'='*70}\n[{key}] {fname}")
        if not fp.exists():
            print(f"  X not found at {fp}"); continue
        try:
            xl = pd.ExcelFile(fp)
        except Exception as e:
            print(f"  X cannot open: {e}"); continue
        print(f"  sheets: {xl.sheet_names}")
        for sh in xl.sheet_names:
            try:
                df = xl.parse(sh, header=0, dtype=str)
            except Exception as e:
                print(f"  - {sh!r}: parse fail {e}"); continue
            amt = _amt_col(df)
            tot = pd.to_numeric(df[amt], errors="coerce").sum() if amt else 0
            print(f"  - {sh!r}: rows={len(df):,} cols={len(df.columns)}  amount_col={amt!r} Σ={tot:,.0f}")
            print(f"      cols: {list(df.columns)[:30]}")
            # f1 capex: show the prebuild filter breakdowns
            if key in ("f1", "f1_capex") and amt:
                for fc in ("Budget Source", "Ledger Hierarchy Level 4"):
                    _dump_break(df, amt, fc)


if __name__ == "__main__":
    main()
