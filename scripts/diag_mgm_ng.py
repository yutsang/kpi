"""Diagnose MGM blank-NG (Section.1) clusters that drive V_OTHER (23=53%, 24=38%). Reads the BUILT
combine files (run build_mgm_24_raw.py first to pick up the new PM→NG1 + project-modal backfill).
Reports, per year:
  - per-Source: rows, Σ|amt|, blank-NG Σ|amt|%  → which tabs still leak blank NG
  - Project_code token overlap WD1↔Capex / leadsheet-keyed↔CAPEX  → can modal-NG fill bridge them?

Run (Windows):  python scripts/diag_mgm_ng.py
Output: prints + results/diag_mgm_ng.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "data" / "mgm" / "raw"
AMT = "Debit minus Credit"
NG = "Section.1"
SRC = "Source"
PC = "Project_code"


def _isblank(s): return s.astype(str).str.strip().isin(["", "nan", "None", "NaN"])


def dump(L, year, fname):
    fp = RAW / fname
    if not fp.exists():
        L.append(f"\n## {year}: X {fp} not found"); return
    df = pd.read_excel(fp, sheet_name="combine", dtype=object)
    a = pd.to_numeric(df[AMT], errors="coerce").fillna(0.0) if AMT in df.columns else pd.Series(0.0, index=df.index)
    atot = a.abs().sum() or 1
    blank = _isblank(df[NG]) if NG in df.columns else pd.Series(False, index=df.index)
    L.append(f"\n## {year}  ({fname})  {len(df):,} rows  Σ={a.sum():,.0f}  "
             f"NG-blank={a.abs()[blank].sum()/atot*100:.1f}% of |amt|")
    if SRC in df.columns:
        L.append(f"   per-{SRC}: rows | Σ|amt|% | blank-NG%-of-source")
        for v, g in df.groupby(df[SRC].astype(str)):
            ga = a.abs()[g.index]; gtot = ga.sum() or 1
            gb = _isblank(g[NG]) if NG in g.columns else pd.Series(False, index=g.index)
            L.append(f"      {str(v)[:18]:18s} {len(g):>7,}  {ga.sum()/atot*100:5.1f}%   blank={ga[gb].sum()/gtot*100:5.1f}%")
    # Project_code token overlap — can a tab's blank rows inherit NG from another tab's filled rows?
    if SRC in df.columns and PC in df.columns:
        filled_pc = set(df.loc[~blank, PC].astype(str).str.strip())
        L.append(f"   blank-NG rows whose Project_code HAS a non-blank-NG sibling (modal-fill can recover):")
        for v, g in df[blank].groupby(df.loc[blank, SRC].astype(str)):
            pc = g[PC].astype(str).str.strip()
            hit = pc.isin(filled_pc)
            ga = a.abs()[g.index]; gtot = ga.sum() or 1
            L.append(f"      {str(v)[:18]:18s} recoverable={ga[hit.values].sum()/gtot*100:5.1f}%  "
                     f"orphan={ga[(~hit).values].sum()/gtot*100:5.1f}%")
        # show a few orphan Project_code tokens to eyeball the namespace mismatch
        orphan_pc = sorted(set(df.loc[blank, PC].astype(str).str.strip()) - filled_pc)[:12]
        filled_sample = sorted(filled_pc)[:12]
        L.append(f"   sample FILLED Project_code: {filled_sample}")
        L.append(f"   sample ORPHAN Project_code (blank-NG, no filled sibling): {orphan_pc}")


def main():
    L = ["# diag_mgm_ng — blank-NG clusters + Project_code token overlap"]
    dump(L, "2023", "mgm_23_raw.xlsx")
    dump(L, "2024", "mgm_24_raw.xlsx")
    out = ROOT / "results" / "diag_mgm_ng.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
