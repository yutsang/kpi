"""inspect_galaxy_adj_cols.py — galaxy 2025 raw Excel 調整金額欄診斷
確認 Combine(clean) sheet 有無 25調整金額/24調整金額/23調整金額 欄
Run: python scripts\inspect_galaxy_adj_cols.py
Out: results\inspect_galaxy_adj_cols.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT  = Path(__file__).resolve().parent.parent
F     = ROOT / "data" / "galaxy" / "raw" / "project_team" / "galaxy_2025.xlsx"
SHEET = "Combine(clean)"
OUT   = ROOT / "results" / "inspect_galaxy_adj_cols.txt"

COMP_ACCTS = {"8107350","7120100","7131000","7141078","7141020",
              "8107500","7141073","7141071","7151020","7132100"}


def main():
    L = ["# inspect_galaxy_adj_cols — galaxy 2025 調整欄診斷", ""]
    if not F.exists():
        # Try alternate paths
        for alt in [ROOT / "data" / "galaxy" / "raw" / "galaxy_2025.xlsx",
                    ROOT / "data" / "galaxy" / "galaxy_2025.xlsx"]:
            if alt.exists():
                F2 = alt; break
        else:
            L.append(f"!! galaxy_2025.xlsx 揾唔到 (tried {F} + alternates)")
            _w(L); return
        # reassign F
        import inspect_galaxy_adj_cols as _self
        _self.F = F2
        _do(F2, L)
    else:
        _do(F, L)
    _w(L)


def _do(fpath, L):
    L.append(f"File: {fpath}")
    xl = pd.ExcelFile(fpath)
    L.append(f"Sheets: {xl.sheet_names}")

    if SHEET not in xl.sheet_names:
        L.append(f"!! '{SHEET}' 唔存在，改用第一個 sheet: {xl.sheet_names[0]}")
        sheet = xl.sheet_names[0]
    else:
        sheet = SHEET

    # Read header-only first for col names
    df_h = pd.read_excel(fpath, sheet_name=sheet, nrows=0)
    cols_all = [str(c).strip() for c in df_h.columns]
    L += ["", "=" * 70, "## 1. 所有欄名"]
    for i, c in enumerate(cols_all):
        L.append(f"  [{i:>3}] {c}")

    # ── 2. 調整相關欄 ──────────────────────────────────────────────────────────
    L += ["", "=" * 70, "## 2. 調整 / 跨年 / amount 相關欄"]
    adj_kw = ["調整", "跨年", "amount", "金額", "mop", "25", "24", "23"]
    adj_cols = [c for c in cols_all if any(k in c for k in adj_kw)]
    L.append(f"  Matched {len(adj_cols)} cols: {adj_cols}")

    # Load only those columns (faster)
    if not adj_cols:
        L.append("  !! 無調整相關欄位找到")
        return

    df = pd.read_excel(fpath, sheet_name=sheet, usecols=adj_cols + [cols_all[0]], dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    n = len(df)
    L.append(f"  Total rows: {n:,}")

    L += ["", "=" * 70, "## 3. 調整/跨年 欄 Σ + non-zero"]
    L.append(f"  {'col':<45} {'Σ':>18}  {'non-zero':>10}")
    for col in adj_cols:
        if col not in df.columns:
            continue
        num = pd.to_numeric(df[col].astype(str).str.replace(",","",regex=False), errors="coerce")
        nz  = int((num.fillna(0) != 0).sum())
        s   = num.sum()
        L.append(f"  {col:<45} {s:>18,.0f}  {nz:>10,}")

    # ── 3. Try to load full file and check comp accounts ──────────────────────
    # Load all cols, try to find account_code col
    df_full = pd.read_excel(fpath, sheet_name=sheet, dtype=str)
    df_full.columns = [str(c).strip() for c in df_full.columns]
    n2 = len(df_full)
    L.append(f"\n  (Full load: {n2:,} rows)")

    # Find account code col
    acct_candidates = [c for c in df_full.columns if any(k in str(c).lower() for k in
                       ["account_code", "account code", "帳號", "acct", "科目"])]
    L += ["", "=" * 70, "## 4. Account code col identification"]
    L.append(f"  Candidates: {acct_candidates}")

    if not acct_candidates:
        # Try first col
        L.append("  Trying col[0] as account code")
        acct_col = df_full.columns[0]
    else:
        acct_col = acct_candidates[0]
    L.append(f"  Using: [{acct_col}]")

    acct_s = df_full[acct_col].astype(str).str.strip()
    comp_mask = acct_s.isin(COMP_ACCTS)
    L.append(f"  Comp account rows: {int(comp_mask.sum()):,} / {n2:,}")

    # For each adj/amount col, show Σ for comp rows vs all rows
    L += ["", "=" * 70, "## 5. 調整/跨年 欄 — comp rows vs all rows Σ"]
    L.append(f"  {'col':<45} {'Σ comp rows':>15}  {'Σ all rows':>15}")
    for col in df_full.columns:
        col_s = str(col)
        if any(k in col_s for k in ["調整", "跨年", "金額", "amount", "mop"]):
            num = pd.to_numeric(df_full[col].astype(str).str.replace(",","",regex=False), errors="coerce").fillna(0)
            sc  = num[comp_mask].sum()
            sa  = num.sum()
            if abs(sa) > 0 or abs(sc) > 0:
                L.append(f"  {col_s:<45} {sc:>15,.0f}  {sa:>15,.0f}")

    # ── Key check: does 25調整金額 Σ(comp) ≈ −31,967萬? ─────────────────────────
    L += ["", "=" * 70, "## 6. 關鍵 check: 25調整金額 comp 行合計 (expected ≈ −31,967萬 = −319,670,000)"]
    for col in df_full.columns:
        if "25調整" in str(col):
            num = pd.to_numeric(df_full[col].astype(str).str.replace(",","",regex=False), errors="coerce").fillna(0)
            sc = num[comp_mask].sum()
            sa = num.sum()
            L.append(f"  [{col}]  Σ comp = {sc:,.0f}  Σ all = {sa:,.0f}")
            L.append(f"         Σ comp 萬 = {sc/1e4:,.2f}  (golden gap = −31,967萬)")
    check_cols = [c for c in df_full.columns if "25調整" in str(c)]
    if not check_cols:
        L.append("  !! 無 '25調整' 欄位")
        L.append(f"  欄名中含 '調整': {[c for c in df_full.columns if '調整' in str(c)]}")

    # ── 7. Sample rows for account 8107350 ──────────────────────────────────────
    L += ["", "=" * 70, "## 7. Account 8107350 sample rows (first 5) — all cols"]
    m8 = acct_s.eq("8107350")
    L.append(f"  8107350 rows: {int(m8.sum()):,}")
    if m8.any():
        sample = df_full[m8].head(5)
        for _, row in sample.iterrows():
            pairs = [f"{c}={str(row[c])[:18]}" for c in df_full.columns
                     if any(k in str(c) for k in ["調整","跨年","金額","amount","Account","人工","分類"])]
            L.append("  " + " | ".join(pairs))


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
