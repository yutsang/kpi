"""Dump the 橫向 / 縱向 sheets from a 投資方向 review Excel into pasteable TSV.

Reads data/review/{ent}_投資方向_{year}.xlsx, extracts sheets 2_橫向 + 3_縱向,
writes them as TSV text files (easy to open + paste back to chat), and prints
the top-N rows by |amount| so you can paste a condensed view directly.

Output:
  data/review/_dump/{ent}_{year}_橫向.tsv
  data/review/_dump/{ent}_{year}_縱向.tsv

Run:
  python scripts/dump_review_tabs.py --entity mgm --year 25
  python scripts/dump_review_tabs.py --entity mgm --year 25 --top 80
  python scripts/dump_review_tabs.py --all          # every file in data/review/
"""
from __future__ import annotations
import argparse, sys, re
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

SHEETS = ["2_橫向", "3_縱向"]


def dump_one(xlsx: Path, top: int):
    out_dir = Path("data/review/_dump")
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = xlsx.stem  # e.g. mgm_投資方向_25
    print(f"\n===== {xlsx.name} =====", flush=True)
    xl = pd.ExcelFile(xlsx)
    for sheet in SHEETS:
        if sheet not in xl.sheet_names:
            print(f"  [skip] sheet '{sheet}' not in {xlsx.name}", flush=True)
            continue
        df = xl.parse(sheet)
        # Full TSV dump
        safe = re.sub(r"[^\w]+", "_", sheet)
        tsv_path = out_dir / f"{stem}__{safe}.tsv"
        df.to_csv(tsv_path, sep="\t", index=False, encoding="utf-8-sig")
        # Condensed top-N by amount
        amt_col = next((c for c in df.columns if "amount" in str(c).lower()), None)
        if amt_col is not None:
            top_df = df.reindex(df[amt_col].abs().sort_values(ascending=False).index).head(top)
        else:
            top_df = df.head(top)
        print(f"\n----- {sheet}  ({len(df):,} rows → {tsv_path}) — top {len(top_df)} by |amount| -----", flush=True)
        print(top_df.to_csv(sep="\t", index=False), flush=True)


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--entity", default=None)
    p.add_argument("--year", default="25")
    p.add_argument("--top", type=int, default=60, help="rows to print per sheet")
    p.add_argument("--all", action="store_true", help="dump every *_投資方向_*.xlsx in data/review/")
    args = p.parse_args()

    review = Path("data/review")
    if args.all:
        files = sorted(review.glob("*_投資方向_*.xlsx"))
    elif args.entity:
        files = [review / f"{args.entity}_投資方向_{args.year}.xlsx"]
    else:
        print("Specify --entity ENT [--year YY] or --all"); sys.exit(1)

    for f in files:
        if not f.exists():
            print(f"❌ {f} missing"); continue
        dump_one(f, args.top)

    print("\n✓ TSV files in data/review/_dump/ — open + paste the rows you want reviewed.", flush=True)


if __name__ == "__main__":
    main()
