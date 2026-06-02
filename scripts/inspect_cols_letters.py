"""Dump an Excel sheet's columns as  Excel-letter | header | n_distinct | samples.

Built for the SJM project-team feedback, which refers to columns by their Excel LETTER
(AN 內部資源, AO 不計入內部資源, AU Net-off, AZ Staff costs, BB 藝人團隊…). This maps each
letter to its real header + shows the value spread, so the classification rules can be
written against the actual column names.

Also used to peek at the Admin Comp summary v2 sheets and the sjm_audit remarks tabs.

Run (list sheet names first if unsure):
  python scripts/inspect_cols_letters.py --file "data/sjm/raw/<SJM raw>.xlsx"
  python scripts/inspect_cols_letters.py --file "data/sjm/raw/<SJM raw>.xlsx" --sheet 表格1
  python scripts/inspect_cols_letters.py --file "data/sjm/raw/Admin Comp summary v2.xlsx" --sheet "combined admin comp"
  python scripts/inspect_cols_letters.py --file "sjm_audit_25.xlsx" --sheet 1_投資方向 --header 0
Outputs results/cols_<file>_<sheet>.tsv
"""
from __future__ import annotations
import argparse, sys, csv, re
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd


def col_letter(idx: int) -> str:
    """0->A, 25->Z, 26->AA, 39->AN, 46->AU, 51->AZ, 53->BB ..."""
    s = ""
    idx += 1
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--file", required=True, help="path to .xlsx (relative to repo root or absolute)")
    p.add_argument("--sheet", default=None, help="sheet name or index; omit to just list sheet names")
    p.add_argument("--header", type=int, default=0, help="header row index (0-based)")
    p.add_argument("--samples", type=int, default=6)
    args = p.parse_args()

    fp = Path(args.file)
    if not fp.exists():
        print(f"X {fp} not found"); sys.exit(1)

    xl = pd.ExcelFile(fp)
    if args.sheet is None:
        print(f"[{fp.name}] sheets:")
        for s in xl.sheet_names:
            print("   ", s)
        print("\n→ re-run with --sheet <name> to dump its columns")
        return

    sheet = args.sheet
    if sheet.isdigit():
        sheet = int(sheet)
    df = xl.parse(sheet, header=args.header, dtype=str)
    print(f"[{fp.name}::{sheet}] {len(df):,} rows × {len(df.columns)} cols  (header_row={args.header})")
    print(f"\n  {'col':<5} {'header':<40} {'n_dist':>7}  samples")
    rows = []
    for i, c in enumerate(df.columns):
        s = df[c].astype("string").fillna("").str.strip()
        nb = s[s.ne("") & s.ne("nan") & s.ne("<NA>")]
        n = nb.nunique()
        samp = " | ".join(str(x)[:24] for x in nb.unique()[:args.samples])
        letter = col_letter(i)
        print(f"  {letter:<5} {str(c)[:40]:<40} {n:>7}  {samp[:70]}")
        rows.append([letter, str(c), n, samp])

    out = Path("results"); out.mkdir(exist_ok=True)
    safe = re.sub(r"[^A-Za-z0-9]+", "_", f"{fp.stem}_{sheet}")[:60]
    with (out / f"cols_{safe}.tsv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter="\t")
        w.writerow(["excel_col", "header", "n_distinct", "samples"])
        w.writerows(rows)
    print(f"\n→ results/cols_{safe}.tsv")


if __name__ == "__main__":
    main()
