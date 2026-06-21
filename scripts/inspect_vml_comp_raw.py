r"""讀 VML raw Excel 直接 dump「Comp類型」breakdown（唔需要 kedro re-run）。

Run（Windows）:
    python scripts\inspect_vml_comp_raw.py
出 results\inspect_vml_comp_raw.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT  = ROOT / "results" / "inspect_vml_comp_raw.txt"

# ── VML 25 (sheet 25JE) ──────────────────────────────────────────────────────
RAW25 = ROOT / "data" / "vml" / "raw" / "vml_2025.xlsx"
# ── VML 23 (sheet Sheet1) ────────────────────────────────────────────────────
RAW23 = ROOT / "data" / "vml" / "raw" / "vml_2023.xlsx"

AMOUNT_CANDIDATES = [
    "調整前金額", "調整前_MOP", "amount", "Amount", "調整前", "pre_amount",
    "adjusted_amount", "金額", "MOP", "amount_mop", "Adjusted Amount",
]

L: list[str] = []


def _find_amount(cols):
    for c in AMOUNT_CANDIDATES:
        if c in cols:
            return c
    # fallback: first col containing "amount" or "金額" case-insensitive
    for c in cols:
        if "amount" in c.lower() or "金額" in c:
            return c
    return None


def _to_wan(s):
    return pd.to_numeric(s, errors="coerce").fillna(0.0) / 10_000


def _inspect(label, raw_path, sheet):
    L.append(f"\n{'='*68}\n## {label}  ({raw_path.name} / {sheet})")
    if not raw_path.exists():
        L.append(f"  !! 找唔到 {raw_path}"); return

    df = pd.read_excel(raw_path, sheet_name=sheet, dtype=str, nrows=None)
    df.columns = [str(c).strip() for c in df.columns]
    L.append(f"  Cols({len(df.columns)}): {list(df.columns)}")
    L.append(f"  Rows: {len(df):,}")

    if "Comp類型" not in df.columns:
        L.append("  !! 冇「Comp類型」欄 — 請確認 sheet 名或欄名")
        # show any comp-related cols
        comp_cols = [c for c in df.columns if "comp" in c.lower() or "贈" in c]
        if comp_cols:
            L.append(f"  Comp-related cols: {comp_cols}")
        return

    amt_col = _find_amount(list(df.columns))
    L.append(f"  Amount col used: {amt_col}")

    df["_ct"] = df["Comp類型"].astype(str).str.strip().replace({"nan": "", "None": ""})
    df["_amt"] = _to_wan(df[amt_col]) if amt_col else pd.Series(0.0, index=df.index)

    total_nonblank = df[df["_ct"] != ""]
    L.append(f"\n  Comp類型 non-blank: {len(total_nonblank):,} 行  Σ={total_nonblank['_amt'].sum():,.0f}萬")

    # unique values
    L.append("\n  ── Comp類型 unique values + 金額 ──")
    g = total_nonblank.groupby("_ct")["_amt"].agg(["sum", "count"]).reset_index()
    g = g.sort_values("sum", ascending=False)
    for _, r in g.iterrows():
        L.append(f"    {r['_ct']:<30}  {r['count']:>6,} 行  Σ={r['sum']:>10,.0f}萬")

    # by account_desc within each comp type (top 5 per type)
    acct_col = next((c for c in df.columns if c.lower() in ("account_desc", "account description", "科目名稱", "account_description")), None)
    if acct_col:
        L.append(f"\n  ── by {acct_col} within each Comp類型 (top 5) ──")
        for ct_val in g["_ct"].tolist():
            sub = total_nonblank[total_nonblank["_ct"] == ct_val]
            sg = sub.groupby(acct_col)["_amt"].sum().sort_values(ascending=False).head(5)
            L.append(f"\n    [{ct_val}]  Σ={sub['_amt'].sum():,.0f}萬")
            for acct, amt in sg.items():
                L.append(f"      {str(acct)[:40]:<40}  {amt:>8,.0f}萬")


def main():
    _inspect("VML 2025", RAW25, "25JE")
    _inspect("VML 2023", RAW23, "Sheet1")

    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
