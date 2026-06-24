"""inspect_galaxy_25_parquet.py — galaxy 25年 parquet 欄名 + comp/staff signal 診斷
讀 data/galaxy/output/company_1_kpi_report.parquet
Run: python scripts\inspect_galaxy_25_parquet.py
Out: results\inspect_galaxy_25_parquet.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
PAR = ROOT / "data" / "galaxy" / "output" / "company_1_kpi_report.parquet"
OUT = ROOT / "results" / "inspect_galaxy_25_parquet.txt"

COMP_KW = ["comp", "room", "fnb", "f&b", "meal", "ticket", "voucher", "compliment",
           "gratis", "complimentary", "贈", "免費", "招待", "comp房", "comp餐"]


def _amt(df, mask):
    for c in ["調整後金額", "amount_mop"]:
        if c in df.columns:
            return pd.to_numeric(df.loc[mask, c], errors="coerce").fillna(0).sum() / 1e4
    return 0.0


def main():
    L = ["# inspect_galaxy_25_parquet — galaxy 25年 parquet 欄覆蓋 + comp signal", ""]
    if not PAR.exists():
        L.append(f"!! 揾唔到 {PAR}"); _w(L); return

    df = pd.read_parquet(PAR)
    mask25 = df["report_period"].astype(str).str.startswith("25")
    d25 = df[mask25]
    n = len(d25)
    L.append(f"galaxy parquet total rows: {len(df):,}  |  25-prefix rows: {n:,}")
    L.append(f"total columns: {len(df.columns)}")

    # ── 1. All column names + non-blank coverage for galaxy 25 rows ─────────────
    L += ["", "=" * 70, "## 1. All columns — non-blank coverage (galaxy 25 rows only)"]
    rows = []
    for col in df.columns:
        try:
            s = d25[col].astype(str).str.strip()
            nb = int(s.ne("").sum())
            pct = nb / n * 100 if n else 0
            rows.append((col, nb, pct))
        except Exception:
            pass
    rows.sort(key=lambda x: -x[1])
    for col, nb, pct in rows:
        marker = "  *** " if pct > 50 else ("  -   " if pct > 5 else "  .   ")
        L.append(f"{marker}{col:<40} {nb:>9,} / {n:,}  ({pct:>5.1f}%)")

    # ── 2. Columns with comp-related keywords in values ──────────────────────────
    L += ["", "=" * 70, "## 2. Columns containing comp-related keywords (galaxy 25 rows)"]
    found_any = False
    for col in df.columns:
        try:
            s = d25[col].astype(str).str.lower()
            m = s.str.contains("|".join(COMP_KW), na=False, regex=True)
            if m.sum() > 0:
                found_any = True
                top = d25.loc[m, col].value_counts().head(8)
                L.append(f"\n  [{col}]  {m.sum():,} rows match comp keywords:")
                for v, c in top.items():
                    L.append(f"    {str(v)[:60]:<62} {c:,}")
        except Exception:
            pass
    if not found_any:
        L.append("  (none found — galaxy 25 has no column with comp keywords)")

    # ── 3. H_LABOR account codes in galaxy 25 (staff overcounting) ──────────────
    L += ["", "=" * 70, "## 3. galaxy 25 H_LABOR (staff) account codes  (golden=14,568萬)"]
    if "horizontal_id" in d25.columns and "account_code" in d25.columns:
        labor = d25[d25["horizontal_id"].eq("H_LABOR")]
        L.append(f"  H_LABOR rows: {len(labor):,}  Σ={_amt(df, mask25 & df['horizontal_id'].eq('H_LABOR')):,.0f}萬  (golden 14,568)")
        if "account_desc" in d25.columns:
            for ac, grp in labor.groupby("account_code"):
                amt = _amt(df, mask25 & df["horizontal_id"].eq("H_LABOR") & df["account_code"].astype(str).eq(str(ac)))
                if amt > 50:
                    desc = grp["account_desc"].mode().iloc[0] if len(grp) else ""
                    L.append(f"    {str(ac):<15} {str(desc)[:45]:<47} rows={len(grp):,}  Σ={amt:,.0f}萬")

    # ── 4. H_HOTEL_ROOM / H_FNB account codes in galaxy 25 (comp sample) ────────
    L += ["", "=" * 70, "## 4. galaxy 25 comp (H_HOTEL_ROOM+H_FNB) account codes  (golden=59,828萬)"]
    COMP_H = {"H_HOTEL_ROOM", "H_VENUE", "H_FNB", "H_COMP_TICKET", "H_COMP_OTHER"}
    if "horizontal_id" in d25.columns and "account_code" in d25.columns:
        comp = d25[d25["horizontal_id"].isin(COMP_H)]
        L.append(f"  comp rows: {len(comp):,}  Σ={_amt(df, mask25 & df['horizontal_id'].isin(COMP_H)):,.0f}萬  (golden 59,828)")
        top_accts = comp["account_code"].value_counts().head(15)
        L.append("  top account_code in comp rows:")
        for ac, cnt in top_accts.items():
            amt = _amt(df, mask25 & df["horizontal_id"].isin(COMP_H) & df["account_code"].astype(str).eq(str(ac)))
            desc = comp.loc[comp["account_code"].astype(str).eq(str(ac)), "account_desc"].mode()
            desc_str = str(desc.iloc[0])[:40] if len(desc) else ""
            L.append(f"    {str(ac):<15} {desc_str:<42} rows={cnt:,}  Σ={amt:,.0f}萬")

    # ── 5. 人工|一級標簽 + 人工|二級標簽 — pt_class_H source ──────────────────
    L += ["", "=" * 70, "## 5. 人工|一級標簽 + 人工|二級標簽 (user says concat = pt_class_H)"]
    col1, col2 = "人工|一級標簽", "人工|二級標簽"
    for col in [col1, col2]:
        if col in d25.columns:
            s = d25[col].astype(str).str.strip()
            nb = s.ne("").sum()
            L.append(f"\n  [{col}]  {nb:,}/{n:,} non-blank ({nb/n*100:.1f}%)  top values:")
            for v, c in s.value_counts().head(20).items():
                L.append(f"    {repr(str(v))[:55]:<57} {c:,}")
        else:
            L.append(f"  [{col}]  !! NOT IN PARQUET — 可能被 step5 drop 或唔存在於 raw")

    # Combo if both found
    if col1 in d25.columns and col2 in d25.columns:
        combo = (d25[col1].astype(str).str.strip() + "|" + d25[col2].astype(str).str.strip()).str.strip("|")
        L.append(f"\n  concat 一級|二級 top25 values (= pt_class_H equivalent):")
        for v, c in combo.value_counts().head(25).items():
            L.append(f"    {str(v)[:60]:<62} {c:,}")
    elif col1 not in d25.columns or col2 not in d25.columns:
        # Try tagged_rows parquet
        tagged = ROOT / "data" / "galaxy" / "interim" / "company_1_tagged_rows.parquet"
        L.append(f"\n  嘗試 tagged_rows.parquet ({tagged.name})...")
        if tagged.exists():
            tr = pd.read_parquet(tagged)
            tr25 = tr[tr.get("report_period", tr.get("year_bucket", pd.Series(dtype=str))).astype(str).str.startswith("25")]
            for col in [col1, col2]:
                if col in tr25.columns:
                    s = tr25[col].astype(str).str.strip()
                    nb = s.ne("").sum()
                    L.append(f"  tagged_rows [{col}]  {nb:,}/{len(tr25):,} non-blank  top5: {dict(s.value_counts().head(5))}")
                else:
                    L.append(f"  tagged_rows [{col}]  NOT FOUND")
        else:
            L.append("  tagged_rows.parquet 亦唔存在")

    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
