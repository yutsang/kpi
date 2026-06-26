r"""inspect_wynn_raw_headers.py — 實證 wynn raw 3 個檔嘅 Entry Voucher 欄名(repr 睇空格)
Run: python scripts\inspect_wynn_raw_headers.py
Out: results\inspect_wynn_raw_headers.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT  = ROOT / "results" / "inspect_wynn_raw_headers.txt"
RAW  = ROOT / "data" / "wynn" / "raw"
# (檔, sheet) — 對齊 conf yearly_sources
FILES = [("wynn_2025.xlsx", "報告投資支出明細賬"), ("wynn_2024.xlsx", 0), ("wynn_2023.xlsx", "Sheet1")]
CONF_AMOUNT = "Entry Voucher Amount/ Expense Amount "   # conf root amount（尾有空格）


def _num(s): return pd.to_numeric(s, errors="coerce")


def main():
    L = ["# wynn raw 3 檔 欄名實證（repr 睇空格）", f"conf root amount = {repr(CONF_AMOUNT)}", ""]
    for fn, sheet in FILES:
        f = RAW / fn
        L += ["", "=" * 70, f"## {fn}  sheet={sheet!r}"]
        if not f.exists():
            L.append(f"  !! 揾唔到 {f}"); continue
        # 只讀 header
        df = pd.read_excel(f, sheet_name=sheet, nrows=0)
        cols = list(df.columns)
        L.append(f"  共 {len(cols)} 欄")
        # 含 amount/voucher/expense/金額 嘅欄，逐個 repr
        amt_like = [c for c in cols if any(k in str(c).lower() for k in
                    ["voucher","expense amount","amount","金額","调整","調整"])]
        L.append("  amount/voucher 相關欄 (repr)：")
        for c in amt_like:
            tail_space = "  ⚠尾有空格" if str(c) != str(c).rstrip() else ""
            L.append(f"     {repr(c)}{tail_space}")
        # conf amount 對唔對得上（連空格 / strip 後）
        exact = CONF_AMOUNT in cols
        stripped_match = [c for c in cols if str(c).strip() == CONF_AMOUNT.strip()]
        L.append(f"  conf amount {repr(CONF_AMOUNT)} 精準喺欄入面? {exact}")
        L.append(f"  strip 後對得上嘅欄: {[repr(c) for c in stripped_match]}")
        # 嗰啲對得上嘅欄 sum（睇值）
        if stripped_match:
            df2 = pd.read_excel(f, sheet_name=sheet, usecols=stripped_match)
            for c in stripped_match:
                L.append(f"     {repr(c)} Σ = {_num(df2[c]).sum()/1e4:,.0f}萬  (非空 {int(_num(df2[c]).notna().sum()):,})")

    OUT.parent.mkdir(exist_ok=True); OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {OUT.relative_to(ROOT)}")

if __name__ == "__main__":
    main()
