r"""inspect_adj_amounts —— 逐家由 adj 算 調整金額 + 調整後金額，按 capex/opex 拆。

user 2026-06-24：請算每家的調整金額和調整後金額按 capex/opex。
逐家讀 adj 標準 tab，揾 amount/調整金額/調整後金額/capex_opex 欄，Σ by capex/opex（萬）。
調整後金額：有 native 欄就用；冇就 = amount + 調整金額。

Run（Windows）:  python scripts\inspect_adj_amounts.py
出 results\inspect_adj_amounts.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import warnings
warnings.filterwarnings("ignore")
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "results" / "inspect_adj_amounts.txt"

# (檔, 候選tab, amount欄候選, 調整金額欄候選, 調整後金額欄候選, capex/opex欄候選)
ENT = {
    "galaxy": ("galaxy_2025_adj.xlsx", ["Combine(clean)"], ["Reported Amount(MOP)"], ["調整金額"], ["調整後金額"], ["Capex/Opex"]),
    "wynn":   ("wynn_2025_adj.xlsx",   ["報告投資支出明細賬"], ["Entry Voucher Amount/ Expense Amount", "Entry Voucher Amount/ Expense Amount "], ["調整金額"], [], ["Capex / Opex", "Capex/Opex"]),
    "vml":    ("vml_2025_adj.xlsx",    ["投資支出明細賬"], ["MOP Amt"], ["調整金額", "調整數"], ["調整後金額"], ["CAPEX_OPEX"]),
    "melco":  ("melco_2025_adj.xlsx",  ["Data"], ["Amount - Amended"], ["調整金額"], [], ["Ledger Type"]),
    "sjm":    ("sjm_2025_adj.xlsx",    ["表格1", "25"], ["Val/COArea Crcy"], ["調整金額"], [], ["CAPEX/OPEX"]),
    "mgm":    ("mgm_2025.xlsx",        ["data"], ["Debit minus Credit"], ["25_Adj", "24_Adj", "23_Adj"], [], ["Capex_Opex", "Capex/Opex"]),
}


def _num(s):
    return pd.to_numeric(pd.Series(s).astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0.0)


def _pick(cols, cands):
    low = {str(c).strip().lower(): c for c in cols}
    for c in cands:
        if c.strip().lower() in low:
            return low[c.strip().lower()]
    # substring fallback
    for c in cands:
        for cc in cols:
            if c.strip().lower() in str(cc).strip().lower():
                return cc
    return None


def _co(x):
    s = str(x).strip().lower()
    if s.startswith("cap"):
        return "Capex"
    if s.startswith("op"):
        return "Opex"
    return "其他/空"


def main():
    L = ["# inspect_adj_amounts —— 逐家 adj 調整金額/調整後金額 by capex/opex（萬）"]
    for ent, (fn, tabs, amts, adjs, afts, cos) in ENT.items():
        p = ROOT / "data" / ent / "raw" / fn
        L.append(f"\n{'='*66}\n## {ent}  ({fn})")
        if not p.exists():
            L.append(f"   !! 揾唔到 {fn}"); continue
        try:
            avail = pd.ExcelFile(p).sheet_names
            tab = next((t for t in tabs if t in avail), None)
            if tab is None:
                L.append(f"   ⚠ tab {tabs} 唔喺；sheets:{avail[:20]}"); continue
            df = pd.read_excel(p, sheet_name=tab, dtype=object)
        except Exception as e:
            L.append(f"   !! 讀失敗：{e}"); continue
        cols = list(df.columns)
        amtc = _pick(cols, amts); aftc = _pick(cols, afts) if afts else None; coc = _pick(cols, cos)
        adjcs = [c for c in (_pick(cols, [a]) for a in adjs) if c is not None]   # 多欄加總（mgm 3 個 _Adj）
        L.append(f"   tab={tab} 行={len(df):,} | amount={amtc!r} 調整={adjcs} 調整後={aftc!r} capex={coc!r}")
        amt = _num(df[amtc]) / 1e4 if amtc else pd.Series(0.0, index=df.index)
        adj = pd.Series(0.0, index=df.index)
        for _c in adjcs:
            adj = adj + _num(df[_c]) / 1e4
        aft = _num(df[aftc]) / 1e4 if aftc else (amt + adj)   # 冇 native 調整後 → amount+調整
        co = df[coc].map(_co) if coc else pd.Series("其他/空", index=df.index)
        g = pd.DataFrame({"co": co, "amount": amt, "調整": adj, "調整後": aft})
        L.append(f"   ── by capex/opex（萬）──")
        L.append(f"      {'':<10}{'amount(報告)':>14}{'調整金額':>13}{'調整後金額':>14}{'行':>9}")
        for k in ["Capex", "Opex", "其他/空"]:
            s = g[g["co"].eq(k)]
            if len(s) == 0:
                continue
            L.append(f"      {k:<10}{s['amount'].sum():>14,.0f}{s['調整'].sum():>13,.0f}{s['調整後'].sum():>14,.0f}{len(s):>9,}")
        L.append(f"      {'總計':<10}{g['amount'].sum():>14,.0f}{g['調整'].sum():>13,.0f}{g['調整後'].sum():>14,.0f}{len(g):>9,}")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
