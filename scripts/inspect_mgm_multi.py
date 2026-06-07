"""Inspect MGM 2023 / 2024 multi-file raw to see if it can be merged into the 2025 schema.

MGM 2025 = a hand-merged single file (mgm_25_raw.xlsx) with sheets 'combine' + 'adjustment',
columns mapping to:  amount='Debit minus Credit' · capex_opex/wd='Source'(CAPEX/WD1..WD5_Patron/
Gaming_OPEX/ADJUSTMENT) · project='Project_code'(項目XXX) · ng11='Section.1'(NG中文) ·
account_code='Ledger Account' · account_desc='Ledger Hierarchy Level 5'  + a golden TSV that
ties the totals + a Master tab for the gaming:(non):clearing ratio split.

23/24 arrive as MULTIPLE files in data/mgm/raw/2023|2024/. This globs them and, per file/sheet,
dumps: rows/cols, the column list, every amount-col Σ (錢喺邊), a head-3 sample, and which raw
column maps to each 2025 slot — so we can design an auto-merge that reproduces the 2025 schema.

Run (Windows):
  python scripts/inspect_mgm_multi.py                 # both 2023 + 2024
  python scripts/inspect_mgm_multi.py --year 2024
Output: prints + results/mgm_multi_inspect.txt
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent

# 2025 target slot → candidate raw-column keywords
SLOTS = {
    "amount 金額":    ["debit minus credit", "debit", "credit", "amount", "金額", "金额", "本位", "translation", "mop"],
    "Source/WD":      ["source", "scenario", "wd", "workday", "capex", "項目組別", "類型", "budget source"],
    "Project_code":   ["project_code", "project code", "project plan task", "項目編", "项目编", "wbs"],
    "Project_name":   ["project_name", "project name", "項目名", "项目名", "投資項目"],
    "Section/NG":     ["section", "項目性質", "项目性质", "nature", "投資領域", "領域", "ng", "category"],
    "Ledger Account": ["ledger account", "account code", "科目代", "gl account", "entry account"],
    "Ledger Hier.":   ["ledger hierarchy", "hierarchy", "level 5", "level 4", "account desc", "科目名"],
    "ratio/Master":   ["ratio", "master", "gaming", "non-gaming", "non gaming", "拆分", "clearing", "%"],
    "year/period":    ["year", "period", "month", "報告", "report", "年度"],
}
AMT_KW = ["debit minus credit", "debit", "credit", "amount", "金額", "金额", "本位", "translation", "投資", "total"]


def auto_header(fp, sheet):
    raw = pd.read_excel(fp, sheet_name=sheet, header=None, nrows=8, dtype=object)
    return max(range(min(6, len(raw))), key=lambda i: raw.iloc[i].notna().sum()) if len(raw) else 0


def dump_file(L, fp):
    L.append(f"\n{'#'*78}\n# FILE: {fp.name}")
    try:
        xl = pd.ExcelFile(fp)
    except Exception as e:
        L.append(f"  X open failed: {e}"); return
    L.append(f"  sheets = {xl.sheet_names}")
    for sheet in xl.sheet_names:
        try:
            hdr = auto_header(fp, sheet)
            df = pd.read_excel(fp, sheet_name=sheet, header=hdr, dtype=object)
        except Exception as e:
            L.append(f"\n  == sheet={sheet!r}: read failed {e}"); continue
        if not len(df):
            L.append(f"\n  == sheet={sheet!r}: empty"); continue
        L.append(f"\n  {'='*68}\n  == sheet={sheet!r}  header_row={hdr}  rows={len(df):,}  cols={len(df.columns)}")
        L.append("  columns: " + " | ".join(str(c) for c in df.columns))
        L.append("  -- 2025 slot 候選欄 --")
        for slot, kws in SLOTS.items():
            cand = [str(c) for c in df.columns if any(k in str(c).lower() for k in kws)]
            if cand: L.append(f"    {slot:15s}: {' | '.join(cand[:6])}")
        amt_cands = [c for c in df.columns if any(k in str(c).lower() for k in AMT_KW)]
        if amt_cands:
            L.append("  -- amount Σ (錢喺邊) --")
            for c in amt_cands[:8]:
                s = pd.to_numeric(df[c], errors="coerce")
                if s.notna().sum(): L.append(f"    {str(c)[:38]:38s}: Σ={s.sum():>18,.0f}  (n={s.notna().sum():,})")
        # distributions for Source/Section-like columns
        for c in df.columns:
            cl = str(c).lower()
            if any(k in cl for k in ["source", "section", "項目性質", "nature", "wd"]):
                vc = df[c].astype(str).str.strip().replace({"nan": "(空)"}).value_counts().head(15)
                if len(vc):
                    L.append(f"    === {c!r} ({df[c].nunique()} distinct) ===")
                    for v, n in vc.items(): L.append(f"        {n:>7,}  {str(v)[:46]}")
        L.append("  -- 頭 3 行 (前 10 欄) --")
        for _, r in df.head(3).iterrows():
            L.append("    " + " | ".join(f"{str(c)[:12]}={str(r[c])[:18]}" for c in df.columns[:10]))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", default="both", help="2023 / 2024 / both")
    a = ap.parse_args()
    years = ["2023", "2024"] if a.year == "both" else [a.year]
    L = ["# MGM 23/24 multi-file inspect — target = 2025 merged schema",
         "# 2025 cols: Debit minus Credit | Source(CAPEX/WD1-5/Gaming_OPEX/ADJUSTMENT) | Project_code | Section.1(NG) | Ledger Account | Ledger Hierarchy Level 5"]
    base = ROOT / "data" / "mgm" / "raw"
    for y in years:
        d = base / y
        files = sorted(p for p in d.glob("*") if p.suffix.lower() in (".xlsx", ".xls", ".xlsm") and not p.name.startswith("~"))
        L.append(f"\n{'='*78}\n=== YEAR {y}  ({d}) — {len(files)} files ===")
        if not files:
            L.append(f"  X no xlsx under {d} (also try: data/mgm/raw/{y}/)")
        for fp in files:
            dump_file(L, fp)
    out = ROOT / "results" / "mgm_multi_inspect.txt"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text("\n".join(str(x) for x in L), encoding="utf-8")
    print("\n".join(str(x) for x in L))
    print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
