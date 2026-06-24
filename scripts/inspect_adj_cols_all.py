"""inspect_adj_cols_all.py — 各 entity tagged_rows parquet 調整欄實際名稱診斷
Run: python scripts\inspect_adj_cols_all.py
Out: results\inspect_adj_cols_all.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT  = ROOT / "results" / "inspect_adj_cols_all.txt"

ENTITIES = [
    ("company_1", "galaxy"),
    ("company_2", "sjm"),
    ("company_3", "wynn"),
    ("company_4", "vml"),
    ("company_5", "melco"),
    ("company_6", "mgm"),
]

ADJ_KEYWORDS = ["調整", "adjust", "adjusted", "adjustment", "前", "後", "lv1", "lv2",
                 "一級", "二級", "level", "2023年度", "2024年度", "2025年度",
                 "初步識別", "類型", "事項", "類別", "備註", "AMT", "Amt"]


def main():
    L = ["# inspect_adj_cols_all — 各 entity 調整相關欄位診斷", ""]

    for comp, alias in ENTITIES:
        pq = ROOT / "data" / alias / "interim" / f"{comp}_tagged_rows.parquet"
        L += ["", "=" * 70, f"## {alias} ({comp})  —  {pq.name}"]
        if not pq.exists():
            L.append(f"  !! parquet 揾唔到: {pq}"); continue

        df = pd.read_parquet(pq, columns=None)
        all_cols = list(df.columns)
        L.append(f"  Total cols: {len(all_cols)}  rows: {len(df):,}")

        # Adjustment-related columns
        adj_cols = [c for c in all_cols if any(k.lower() in str(c).lower() for k in ADJ_KEYWORDS)]
        L.append(f"  Adjustment-related cols ({len(adj_cols)}):")
        for c in adj_cols:
            col_data = df[c]
            non_null = int(col_data.notna().sum())
            try:
                num = pd.to_numeric(col_data, errors="coerce")
                s = num.sum()
                non_zero = int((num.fillna(0) != 0).sum())
                L.append(f"    [{c}]  non-null={non_null:,}  numeric_Σ={s:,.0f}  non_zero={non_zero:,}")
            except Exception:
                vc = col_data.dropna().astype(str).value_counts().head(4)
                L.append(f"    [{c}]  non-null={non_null:,}  top: " + " | ".join(f"{v}:{n}" for v, n in vc.items()))

        # year_bucket breakdown for key amount cols
        if "year_bucket" in df.columns:
            L.append(f"\n  year_bucket dist:")
            vc = df["year_bucket"].astype(str).value_counts().sort_index()
            for b, n in vc.items():
                L.append(f"    {b}: {n:,} rows")

        # amount_mop Σ by year_bucket
        if "amount_mop" in df.columns and "year_bucket" in df.columns:
            L.append(f"\n  amount_mop Σ by year_bucket:")
            num_amt = pd.to_numeric(df["amount_mop"], errors="coerce").fillna(0)
            for b in sorted(df["year_bucket"].astype(str).unique()):
                m = df["year_bucket"].astype(str).eq(b)
                L.append(f"    {b}: {num_amt[m].sum()/1e4:,.0f}萬")

    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
