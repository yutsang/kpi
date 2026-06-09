"""MGM 2025 — the 'combine' sheet (what our pipeline reads) has WD1 430120 = 23.8M / 614 rows,
but the 'NGO_WD1' source sheet (what you filtered) = 12.59M. So `combine` itself carries ~2×.
This scans EVERY sheet in mgm_25_raw.xlsx for 'Ledger Account' contains <needle>, reporting
rows + SUM (and WD1-only if a Source col exists), so we see which source sheet(s) the combine
rows were stacked from — i.e. is combine = NGO_WD1 + another tab that repeats the same comp?

Run (Windows):  python scripts/inspect_mgm_25_origin.py 430120   (no arg → 430120)
Output: prints + results/inspect_mgm_25_origin.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "data" / "mgm" / "raw" / "mgm_25_raw.xlsx"
ACC, SRC, AMT = "Ledger Account", "Source", "Debit minus Credit"


def _col(df, *names):
    for name in names:
        if name in df.columns:
            return name
        for c in df.columns:
            if str(c).strip() == name:
                return c
    return None


def main():
    needle = sys.argv[1] if len(sys.argv) > 1 else "430120"
    L = [f"# inspect_mgm_25_origin — scan ALL sheets for '{ACC}' contains '{needle}' (find the 2× source)"]
    if not RAW.exists():
        L.append(f"X {RAW} missing"); _w(L); return
    xls = pd.ExcelFile(RAW)
    L.append(f"sheets: {xls.sheet_names}\n")
    L.append(f"{'sheet':22s} {'rows':>7s}  {'all-Σ(M)':>10s} {'all-n':>6s}   {'WD1-Σ(M)':>10s} {'WD1-n':>6s}")
    for sh in xls.sheet_names:
        try:
            df = pd.read_excel(RAW, sheet_name=sh)
        except Exception as e:
            L.append(f"{sh:22s}  (read error: {e})"); continue
        ac = _col(df, ACC)
        am = _col(df, AMT)
        if not ac or not am:
            L.append(f"{sh:22s} {len(df):>7,}  (no Ledger Account / Amt col)"); continue
        hit = df[ac].astype("string").fillna("").str.contains(needle, regex=False, na=False)
        amt = pd.to_numeric(df[am], errors="coerce").fillna(0.0)
        all_s, all_n = amt[hit].sum(), int(hit.sum())
        sc = _col(df, SRC)
        if sc is not None:
            wd1 = hit & df[sc].astype("string").fillna("").str.strip().eq("WD1")
            w_s, w_n = amt[wd1].sum(), int(wd1.sum())
            wtxt = f"{w_s/1e6:>10.3f} {w_n:>6,}"
        else:
            wtxt = f"{'(no Source)':>17s}"
        L.append(f"{sh:22s} {len(df):>7,}  {all_s/1e6:>10.3f} {all_n:>6,}   {wtxt}")
    L.append("\nReading: the sheet(s) whose WD1-Σ add up to combine's 23.8M are the merge sources.")
    L.append("If e.g. NGO_WD1=12.59M and project/Project_all=11.2M → combine stacked both = double-count.")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_mgm_25_origin.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
