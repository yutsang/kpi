"""inspect_sjm_raw.py — sjm_2025.xlsx 'data' tab 結構診斷
目的：睇 'data' tab 有哪些列名，用於改 conf bucket_amount_cols + capex 欄
Run: python scripts\inspect_sjm_raw.py
Out: results\inspect_sjm_raw.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT  = ROOT / "results" / "inspect_sjm_raw.txt"

# Search common paths for sjm_2025.xlsx
CANDIDATES = [
    ROOT / "data" / "sjm" / "raw" / "sjm_2025.xlsx",
    ROOT / "data" / "sjm" / "sjm_2025.xlsx",
    ROOT / "sjm_2025.xlsx",
]


def _find_xls():
    for p in CANDIDATES:
        if p.exists():
            return p
    # Broader search
    for p in ROOT.rglob("sjm_2025.xlsx"):
        return p
    return None


def main():
    L = ["# inspect_sjm_raw — sjm_2025.xlsx 'data' tab 結構", ""]

    xls = _find_xls()
    if xls is None:
        L.append(f"!! sjm_2025.xlsx 揾唔到，試過: {[str(c) for c in CANDIDATES]}")
        _w(L); return

    L.append(f"File: {xls}")

    # List all sheets
    xl = pd.ExcelFile(xls)
    L.append(f"Sheets: {xl.sheet_names}")
    L.append("")

    # Read 'data' tab (first 5000 rows for speed)
    if "data" not in xl.sheet_names:
        L.append("!! 'data' sheet 唔存在！")
        L.append("  Available sheets: " + str(xl.sheet_names))
        _w(L); return

    df = pd.read_excel(xls, sheet_name="data", header=0, nrows=10000)
    n = len(df)
    L.append(f"'data' tab preview: {n:,} rows (10k cap), {len(df.columns)} cols")
    L.append("")

    # ── 1. All columns with coverage ────────────────────────────────────────────
    L += ["=" * 70, "## 1. All column names + coverage + top values (data tab)"]
    for i, col in enumerate(df.columns):
        s = df[col].astype(str).str.strip()
        nb = int(s.ne("").sum())
        pct = nb / n * 100 if n else 0
        top = s[s.ne("") & s.ne("nan") & s.ne("None")].value_counts().head(4)
        top_str = " | ".join(f"{repr(str(v)[:18])}:{c}" for v, c in top.items())
        L.append(f"  [{i:>3}] {str(col)[:50]:<52} {nb:>6}/{n} ({pct:>4.0f}%) {top_str[:90]}")

    # ── 2. Amount-like columns (numeric, non-zero) ───────────────────────────────
    L += ["", "=" * 70, "## 2. Numeric amount columns (sum > 0)"]
    for col in df.columns:
        try:
            num = pd.to_numeric(df[col], errors="coerce")
            if num.notna().sum() > n * 0.1 and num.sum() != 0:
                L.append(f"  {str(col)[:50]:<52}  Σ={num.sum():>15,.0f}  non-null={num.notna().sum():,}")
        except Exception:
            pass

    # ── 3. Capex/Opex-like columns ───────────────────────────────────────────────
    L += ["", "=" * 70, "## 3. Capex/Opex indicator columns"]
    cap_kw = ["capex", "opex", "資本", "費用", "支出性質", "報告"]
    for col in df.columns:
        col_l = str(col).lower()
        if any(k in col_l for k in cap_kw):
            s = df[col].astype(str).str.strip()
            vc = s[s.ne("") & s.ne("nan")].value_counts().head(8)
            L.append(f"\n  [{col}]")
            for v, c in vc.items():
                L.append(f"    {repr(str(v))[:50]}: {c:,}")

    # ── 4. 調整 / adjustment columns ────────────────────────────────────────────
    L += ["", "=" * 70, "## 4. 調整 / adjustment related columns"]
    adj_kw = ["調整", "adjustment", "adj", "跨年", "lv", "level"]
    for col in df.columns:
        col_l = str(col).lower()
        if any(k in col_l for k in adj_kw):
            s = df[col].astype(str).str.strip()
            nb = int(s.ne("").sum())
            num = pd.to_numeric(df[col], errors="coerce")
            is_numeric = num.notna().sum() > nb * 0.3
            if is_numeric:
                L.append(f"  [NUMERIC] {str(col)[:50]:<52}  Σ={num.sum():>12,.0f}  non-null={num.notna().sum():,}")
            else:
                vc = s[s.ne("") & s.ne("nan")].value_counts().head(5)
                top_str = " | ".join(f"{repr(str(v)[:15])}:{c}" for v, c in vc.items())
                L.append(f"  [TEXT]    {str(col)[:50]:<52}  non-blank={nb:,}  {top_str[:70]}")

    # ── 5. AL-AS columns (user says Y=flag) ─────────────────────────────────────
    L += ["", "=" * 70, "## 5. Columns around position AL-AS (Excel columns 38-45 = index 37-44)"]
    for i, col in enumerate(df.columns):
        if 35 <= i <= 48:
            s = df[col].astype(str).str.strip()
            vc = s[s.ne("") & s.ne("nan")].value_counts().head(5)
            top_str = " | ".join(f"{repr(str(v)[:12])}:{c}" for v, c in vc.items())
            L.append(f"  [col {i:>2} / Excel col {chr(65+i//26-1) if i>=26 else ''}{chr(65+i%26)}] {str(col)[:45]:<47} {top_str[:70]}")

    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
