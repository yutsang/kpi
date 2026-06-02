"""Inspect MGM raw source files + current prebuild output.

Usage (Windows or Mac, repo root):
  python -m kpi.tools.mgm_inspect

Output: a compact stdout report listing
  - Each raw file in data/mgm/raw/  → sheet names + columns + row count
  - For each sheet (sampled): top 5 unique values per column (truncated)
  - data/mgm/interim/company_6_raw.parquet status + per-source row counts

Goal: give enough info for the user to fill in kpi_sheets.xlsx for MGM
      WITHOUT loading every cell. Keep extraction minimal.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

from kpi.lib.conf import ROOT
from kpi.lib.io_setup import force_unbuffered_io

force_unbuffered_io()


RAW_DIR = ROOT / "data" / "mgm" / "raw"
PARQUET = ROOT / "data" / "mgm" / "interim" / "company_6_raw.parquet"

MAX_COL_PRINT = 60          # truncate long column names
MAX_SAMPLES_PER_COL = 5     # unique values to show per column
SAMPLE_ROWS = 200           # rows to read per sheet for sampling (cheap)


def _short(v) -> str:
    s = str(v)
    return s if len(s) <= MAX_COL_PRINT else s[: MAX_COL_PRINT - 1] + "…"


def inspect_xlsx(fpath: Path) -> None:
    print(f"\n── {fpath.name} ──")
    try:
        xl = pd.ExcelFile(fpath, engine="openpyxl")
    except Exception as e:
        print(f"  [error] {e}")
        return
    for sheet in xl.sheet_names:
        try:
            df = pd.read_excel(fpath, sheet_name=sheet, dtype=str, nrows=SAMPLE_ROWS, engine="openpyxl")
        except Exception as e:
            print(f"  [{sheet}] read error: {e}")
            continue
        print(f"  Sheet: {sheet!r}   sampled {len(df)} rows × {len(df.columns)} cols")
        for col in df.columns:
            vals = df[col].dropna().astype(str).str.strip()
            vals = vals[vals != ""]
            uniq = vals.unique()[:MAX_SAMPLES_PER_COL]
            if len(uniq) == 0:
                continue
            samples = " | ".join(_short(v) for v in uniq)
            print(f"     • {_short(col):<60}  → {samples}")


def inspect_parquet() -> None:
    print(f"\n── current prebuild output: {PARQUET.name} ──")
    if not PARQUET.exists():
        print("  (not found — prebuild has not run on this machine)")
        return
    df = pd.read_parquet(PARQUET)
    print(f"  rows: {len(df):,}   cols: {len(df.columns)}")
    print(f"  columns: {list(df.columns)}")
    print()
    if "source" in df.columns:
        print("  source × wd distribution (row counts):")
        ct = df.groupby(["source", "wd"], dropna=False).size().rename("rows").reset_index()
        for _, r in ct.iterrows():
            print(f"     {r['source']:<14} {str(r['wd']):<14} {r['rows']:>10,}")
    if "section" in df.columns:
        print("\n  section value counts (top 15):")
        for v, c in df["section"].value_counts(dropna=False).head(15).items():
            print(f"     {_short(v):<60} {c:>8,}")
    if "project_name" in df.columns:
        print(f"\n  unique project_name: {df['project_name'].nunique(dropna=False):,}")
        print("  project_name samples (first 10 by row count):")
        for v, c in df["project_name"].value_counts(dropna=False).head(10).items():
            print(f"     {_short(v):<60} {c:>8,}")


def main() -> None:
    print(f"MGM raw dir: {RAW_DIR}")
    if not RAW_DIR.exists():
        print("  (raw dir not found — copy MGM source files to data/mgm/raw/)")
    else:
        files = sorted(p for p in RAW_DIR.iterdir() if p.suffix.lower() in (".xlsx", ".xls"))
        if not files:
            print("  (no xlsx files found)")
        for f in files:
            inspect_xlsx(f)
    inspect_parquet()


if __name__ == "__main__":
    main()
