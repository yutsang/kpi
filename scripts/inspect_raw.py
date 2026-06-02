"""Inspect a raw source xlsx: list sheets, columns, and value distributions for the
columns that matter for wiring a new year (V/H source cols, year-filter col, amount).

Run (Windows):
  python scripts/inspect_raw.py --file vml_2023.xlsx
  python scripts/inspect_raw.py --file vml_2023.xlsx --sheet 23JE
  python scripts/inspect_raw.py --file vml_2023.xlsx --sheet 23JE --col 類別1 分類1
Output: prints + results/<stem>_inspect.txt
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
# columns whose value-distribution we dump by default (V/H source, year, amount, capex)
KW = ["類別", "分類", "年", "year", "yr", "投資", "人工", "性質", "調整後金額",
      "capex", "opex", "金額", "項目", "ng", "category"]


def find_file(name):
    p = Path(name)
    if p.exists(): return p
    hits = list((ROOT / "data").glob(f"*/raw/{name}")) + list((ROOT / "data").glob(f"*/raw/*{name}*"))
    return hits[0] if hits else None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", required=True)
    ap.add_argument("--sheet", default=None)
    ap.add_argument("--col", nargs="*", default=None, help="只 dump 呢啲欄（覆寫關鍵字偵測）")
    ap.add_argument("--top", type=int, default=25)
    args = ap.parse_args()
    fp = find_file(args.file)
    if not fp:
        print(f"X {args.file} not found under data/*/raw/"); return

    xl = pd.ExcelFile(fp)
    lines = [f"# {fp.name}  sheets = {xl.sheet_names}"]
    print(lines[0])
    sheet = args.sheet if args.sheet is not None else xl.sheet_names[0]
    # header auto-detect: scan first 6 rows for the one with most non-null
    raw = pd.read_excel(fp, sheet_name=sheet, header=None, nrows=8, dtype=object)
    hdr = max(range(min(6, len(raw))), key=lambda i: raw.iloc[i].notna().sum())
    df = pd.read_excel(fp, sheet_name=sheet, header=hdr, dtype=object)
    lines.append(f"\n## sheet={sheet!r}  header_row={hdr}  rows={len(df):,}  cols={len(df.columns)}")
    lines.append("columns: " + " | ".join(str(c) for c in df.columns))

    if args.col:
        targets = [c for c in df.columns if any(k.lower() in str(c).lower() for k in args.col)]
    else:
        targets = [c for c in df.columns if any(k.lower() in str(c).lower() for k in KW)]
    for c in targets:
        vc = df[c].astype(str).str.strip().replace({"nan": "(空)"}).value_counts().head(args.top)
        lines.append(f"\n=== {c!r}  ({df[c].notna().sum():,} non-null, {df[c].nunique()} distinct) ===")
        for v, n in vc.items():
            lines.append(f"  {n:>7,}  {str(v)[:60]}")
    out = ROOT / "results" / f"{fp.stem}_inspect.txt"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text("\n".join(lines), encoding="utf-8")
    print("\n".join(lines[1:]))
    print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
