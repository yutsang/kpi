"""inspect_galaxy_25_period.py — galaxy_2025.xlsx Period 欄實際值診斷
Run: python scripts\inspect_galaxy_25_period.py
Out: results\inspect_galaxy_25_period.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
F    = ROOT / "data" / "galaxy" / "raw" / "galaxy_2025.xlsx"
OUT  = ROOT / "results" / "inspect_galaxy_25_period.txt"


def main():
    L = ["# inspect_galaxy_25_period — galaxy 2025 Period 欄診斷", ""]
    if not F.exists():
        L.append(f"!! {F} 揾唔到"); _w(L); return

    df = pd.read_excel(F, sheet_name="Combine(clean)", dtype=str, nrows=5)
    df.columns = [str(c).strip() for c in df.columns]
    L.append(f"Cols (first 30): {list(df.columns[:30])}")

    # Find period-like columns
    L += ["", "## 含 '年' / 'Period' / '報告' / '期後' / 'SY' 嘅欄:"]
    kw = ["年", "Period", "期後", "SY", "報告", "跨年"]
    found = [c for c in df.columns if any(k.lower() in str(c).lower() for k in kw)]
    L.append(f"  Found: {found}")

    # Full load just to check Period values
    df_full = pd.read_excel(F, sheet_name="Combine(clean)", dtype=str,
                            usecols=lambda c: any(k.lower() in str(c).strip().lower()
                                                  for k in kw + ["Period"]))
    df_full.columns = [str(c).strip() for c in df_full.columns]
    L.append(f"\nRows loaded: {len(df_full):,}")

    for c in df_full.columns:
        vc = df_full[c].astype(str).str.strip().value_counts(dropna=False).head(10)
        total = len(df_full)
        L += [f"\n## [{c}] value counts:"]
        for v, cnt in vc.items():
            L.append(f"  {repr(str(v)):<45}  {cnt:>7,}  ({cnt/total*100:.1f}%)")

    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
