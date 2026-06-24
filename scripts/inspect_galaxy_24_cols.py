"""inspect_galaxy_24_cols.py — galaxy_2024.xlsx 欄位結構 + year/Period 欄診斷
Run: python scripts\inspect_galaxy_24_cols.py
Out: results\inspect_galaxy_24_cols.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT  = ROOT / "results" / "inspect_galaxy_24_cols.txt"

FILES = [
    ROOT / "data" / "galaxy" / "raw" / "galaxy_2024.xlsx",
    ROOT / "data" / "galaxy" / "raw" / "galaxy_2024_23sy.xlsx",
]


def _num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s.astype(str).str.replace(",","",regex=False), errors="coerce").fillna(0.0)


def inspect_file(path: Path, L: list):
    L += ["", "=" * 70, f"## FILE: {path.name}"]
    if not path.exists():
        L.append(f"  !! 揾唔到 ({path})"); return

    # Try to find sheets
    try:
        xl = pd.ExcelFile(path)
        L.append(f"  Sheets: {xl.sheet_names}")
        sheet = xl.sheet_names[0]
    except Exception as e:
        L.append(f"  !! Error opening: {e}"); return

    try:
        df = pd.read_excel(path, sheet_name=sheet, dtype=str, nrows=5)
        df.columns = [str(c).strip() for c in df.columns]
        L.append(f"  Loaded {len(df)} sample rows, {len(df.columns)} cols from sheet [{sheet}]")
    except Exception as e:
        L.append(f"  !! Error reading: {e}"); return

    # Full load for Σ
    df_full = pd.read_excel(path, sheet_name=sheet, dtype=str)
    df_full.columns = [str(c).strip() for c in df_full.columns]
    n = len(df_full)
    L.append(f"  Full: {n:,} rows")

    # All columns
    L += ["", "  ALL columns:"]
    for i, c in enumerate(df_full.columns):
        L.append(f"    [{i:>3}] {c}")

    # Look for year / period / 年度 columns
    L += ["", "  Columns containing '年', 'Period', 'year', '期後', 'SY', '報告':"]
    keywords = ["年", "Period", "year", "期後", "SY", "報告", "跨年", "調整"]
    found = [c for c in df_full.columns if any(k.lower() in str(c).lower() for k in keywords)]
    for c in found:
        vc = df_full[c].astype(str).str.strip().value_counts(dropna=False).head(8)
        L.append(f"    [{c}] → " + " | ".join(f"{v}:{cnt}" for v, cnt in vc.items()))

    # Look for amount-like columns
    L += ["", "  Numeric cols (Σ, top 15 by abs):"]
    num_sums = {}
    for c in df_full.columns:
        s = _num(df_full[c])
        if s.abs().sum() > 0:
            num_sums[c] = s.sum()
    sorted_cols = sorted(num_sums.items(), key=lambda x: abs(x[1]), reverse=True)[:15]
    for c, s in sorted_cols:
        L.append(f"    {str(c):<50}  Σ={s:>20,.0f}")

    # Sample rows (first 3)
    L += ["", "  First 3 rows (key cols):"]
    show_cols = found[:6] + [c for c, _ in sorted_cols[:5] if c not in found][:5]
    show_cols = list(dict.fromkeys(show_cols))[:10]
    for _, row in df_full.head(3).iterrows():
        L.append("    " + " | ".join(f"{str(c)[:15]}={str(row[c])[:20]}" for c in show_cols))


def main():
    L = ["# inspect_galaxy_24_cols — galaxy 2024 raw 欄位結構診斷", ""]
    for f in FILES:
        inspect_file(f, L)

    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
