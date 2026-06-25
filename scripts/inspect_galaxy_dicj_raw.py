r"""inspect_galaxy_dicj_raw.py — 證 galaxy 2025 raw 有「DICJ Code」但 conf 冇讀
Run: python scripts\inspect_galaxy_dicj_raw.py
Out: results\inspect_galaxy_dicj_raw.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT  = ROOT / "results" / "inspect_galaxy_dicj_raw.txt"
RAW  = ROOT / "data" / "galaxy" / "raw"


def _blank(s): return s.fillna("").astype(str).str.strip().isin(["", "nan", "None", "NaN", "<NA>"])


def main():
    L = ["# galaxy raw DICJ 欄診斷", ""]

    # 2025
    f25 = RAW / "galaxy_2025.xlsx"
    L += ["=" * 60, f"## 2025  {f25.name}  sheet=Combine(clean)"]
    if f25.exists():
        df = pd.read_excel(f25, sheet_name="Combine(clean)", dtype=str)
        df.columns = [str(c).strip() for c in df.columns]
        dcol = next((c for c in df.columns if c.lower().replace(" ", "") in ("dicjcode", "dicj_code")), None)
        L.append(f"  rows={len(df):,}  DICJ 欄名={dcol!r}")
        if dcol:
            nb = (~_blank(df[dcol])).sum()
            L.append(f"  「{dcol}」非空={nb:,}/{len(df):,} ({nb/len(df)*100:.1f}%)  空={len(df)-nb:,}")
            L.append(f"  非空樣本: {df.loc[~_blank(df[dcol]), dcol].dropna().unique()[:15].tolist()}")
            if "Period" in df.columns:
                L.append("  by Period 非空率:")
                for p, g in df.groupby(df["Period"].fillna("")):
                    nb2 = (~_blank(g[dcol])).sum()
                    L.append(f"     {p or '<空>':<16} {nb2:>7,}/{len(g):>7,} 非空 ({nb2/len(g)*100:.0f}%)")
        L.append(f"  有冇 lowercase 'dicj_code' 欄: {'dicj_code' in df.columns}")

    # 2024 / 2023
    for yr in ("2024", "2023"):
        f = RAW / f"galaxy_{yr}.xlsx"
        L += ["", "=" * 60, f"## {yr}  {f.name}  sheet=data"]
        if f.exists():
            df = pd.read_excel(f, sheet_name="data", dtype=str, nrows=50000)
            df.columns = [str(c).strip() for c in df.columns]
            has = "dicj_code" in df.columns
            L.append(f"  rows(sample)={len(df):,}  有 'dicj_code' 欄={has}")
            if has:
                nb = (~_blank(df["dicj_code"])).sum()
                L.append(f"  dicj_code 非空={nb:,}/{len(df):,}  樣本={df.loc[~_blank(df['dicj_code']),'dicj_code'].dropna().unique()[:10].tolist()}")
            dcap = [c for c in df.columns if "dicj" in c.lower()]
            L.append(f"  含 'dicj' 嘅欄: {dcap}")

    OUT.parent.mkdir(exist_ok=True); OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {OUT.relative_to(ROOT)}")

if __name__ == "__main__":
    main()
