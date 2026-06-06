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

# wiring slots → candidate-column keywords (printed as "wiring 候選欄" per sheet)
SLOTS = {
    "amount 金額":  ["金額", "amount", "調整後", "val/", "crcy", "debit", "credit", "本位幣"],
    "project 項目": ["項目名", "项目名", "project name", "name of investment", "投資項目", "项目", "project code"],
    "subproject":   ["subproject", "sub project", "sub-project", "子項目", "initiative", "项目名称"],
    "account_code": ["account code", "科目代", "cost element", "gl account", "ledger account", "會計科目", "entry account"],
    "account_desc": ["account desc", "科目名", "cost element desc", "ledger hierarchy", "gl account desc", "科目摘要"],
    "capex/opex":   ["capex", "opex", "資本", "費用性質"],
    "NG/性質 分類": ["項目性質", "項目類型", "項目分類", "範疇", "ng11", "ng category", "投資領域", "投資方向", "nature", "分類", "類別"],
    "unique_id":    ["唯一", "unique id", "識別碼", "uid", "unique_id"],
    "vendor":       ["vendor", "供應商", "供应商", "廠商", "supplier"],
    "desc 摘要":    ["description", "摘要", "memo", "narration", "journal line"],
    "year 年份":    ["是否2", "年度", "year", " yr", "report"],
}


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
    sheets = [args.sheet] if args.sheet is not None else list(xl.sheet_names)
    for sheet in sheets:
        # header auto-detect: scan first 6 rows for the one with most non-null
        raw = pd.read_excel(fp, sheet_name=sheet, header=None, nrows=8, dtype=object)
        hdr = max(range(min(6, len(raw))), key=lambda i: raw.iloc[i].notna().sum()) if len(raw) else 0
        df = pd.read_excel(fp, sheet_name=sheet, header=hdr, dtype=object)
        lines.append(f"\n{'='*72}\n## sheet={sheet!r}  header_row={hdr}  rows={len(df):,}  cols={len(df.columns)}")
        lines.append("columns: " + " | ".join(str(c) for c in df.columns))

        # wiring 候選欄 — which raw column maps to each conf slot
        lines.append("\n-- wiring 候選欄 (raw col → conf slot) --")
        for slot, kws in SLOTS.items():
            cand = [str(c) for c in df.columns if any(k.lower() in str(c).lower() for k in kws)]
            if cand:
                lines.append(f"  {slot:14s}: {' | '.join(cand[:6])}")

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
    print("\n".join(lines))
    print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
