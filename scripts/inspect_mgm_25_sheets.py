"""MGM 2025 — our table WD1 430120 = 23.8M but raw filter = 12.59M (~2× WITHIN year 25).
That rules out year/sign/source → smells like double-count. step0 reads BOTH sheets
(raw_sheet: ["combine", "adjustment"]). This reads the raw xlsx and splits WD1+<needle> by
sheet, then checks for duplicate rows (in-sheet) and cross-sheet overlap, to pinpoint the 2×.

Run (Windows):  python scripts/inspect_mgm_25_sheets.py 430120   (no arg → 430120)
Output: prints + results/inspect_mgm_25_sheets.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "data" / "mgm" / "raw" / "mgm_25_raw.xlsx"
SHEETS = ["combine", "adjustment"]
SRC, ACC, AMT = "Source", "Ledger Account", "Debit minus Credit"


def _col(df, name):
    if name in df.columns: return name
    for c in df.columns:
        if str(c).strip() == name: return c
    return None


def main():
    needle = sys.argv[1] if len(sys.argv) > 1 else "430120"
    L = [f"# inspect_mgm_25_sheets — WD1 '{ACC}' contains '{needle}' split by sheet (find the 2×)"]
    if not RAW.exists():
        L.append(f"X {RAW} missing"); _w(L); return
    xls = pd.ExcelFile(RAW)
    L.append(f"sheets in file: {xls.sheet_names}")
    subs = {}
    for sh in SHEETS:
        if sh not in xls.sheet_names:
            L.append(f"\n## sheet '{sh}' NOT in file — skipped"); continue
        df = pd.read_excel(RAW, sheet_name=sh)
        sc, ac, am = _col(df, SRC), _col(df, ACC), _col(df, AMT)
        L.append(f"\n## sheet '{sh}': rows={len(df):,}  Source={sc!r} Ledger={ac!r} Amt={am!r}")
        if not (sc and ac and am):
            L.append("   !! missing one of Source/Ledger/Amt — cannot filter"); continue
        wd1 = df[sc].astype("string").fillna("").str.strip().eq("WD1")
        hit = df[ac].astype("string").fillna("").str.contains(needle, regex=False, na=False)
        sub = df[wd1 & hit].copy()
        amt = pd.to_numeric(sub[am], errors="coerce").fillna(0.0)
        L.append(f"   WD1 & contains {needle}: rows={len(sub):,}  SUM={amt.sum():,.0f}  |abs|={amt.abs().sum():,.0f}")
        # in-sheet duplicate full rows
        dfull = sub.duplicated(keep=False).sum()
        L.append(f"   exact-duplicate rows within sheet: {dfull}")
        subs[sh] = sub.reset_index(drop=True)

    # cross-sheet overlap: are combine rows also present in adjustment (same values)?
    if len(subs) == 2:
        a, b = subs["combine"], subs["adjustment"]
        common = [c for c in a.columns if c in b.columns]
        ka = a[common].astype("string").fillna("").agg("|".join, axis=1)
        kb = b[common].astype("string").fillna("").agg("|".join, axis=1)
        inter = set(ka) & set(kb)
        L.append(f"\n## cross-sheet overlap (combine ∩ adjustment on {len(common)} shared cols):")
        L.append(f"   combine keys={len(set(ka))}  adjustment keys={len(set(kb))}  identical-row keys in BOTH={len(inter)}")
        if inter:
            ov = pd.to_numeric(b.loc[kb.isin(inter), _col(b, AMT)], errors="coerce").fillna(0.0)
            L.append(f"   adjustment amount that duplicates combine: {ov.sum():,.0f}  ({len(ov)} rows)")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_mgm_25_sheets.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
