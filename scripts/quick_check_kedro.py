r"""quick_check_kedro —— 唔跑 prep_tableau，直接讀各家 kedro 輸出(kpi_report.parquet)
算 報告 capex/opex by report_period(25/24/23 prefix)，對 golden(重分類後)。

user 2026-06-24：等 6 家 ready 先 prep 太耐，想先大約 check 換檔後啱唔啱。
galaxy capex 正規化：final_capex_opex == "Capex" 先 Capex，其他 → Opex（人工|一級標簽）。

Run（Windows）:  python scripts\quick_check_kedro.py
出 results\quick_check_kedro.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "results" / "quick_check_kedro.txt"
ENTS = {"galaxy": 1, "sjm": 2, "wynn": 3, "vml": 4, "melco": 5, "mgm": 6}
# golden 報告(重分類後) 萬：capex / opex
GCAP = {"galaxy": 131838, "wynn": 141056, "vml": 135760, "melco": 127111, "mgm": 155597, "sjm": 157955}
GOPX = {"galaxy": 213555, "wynn": 74294, "vml": 116590, "melco": 129484, "mgm": 100336, "sjm": 38873}


def _num(s):
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def _pick(cols, cands):
    low = {str(c).strip().lower(): c for c in cols}
    for c in cands:
        if c.lower() in low:
            return low[c.lower()]
    for c in cands:
        for cc in cols:
            if c.lower() in str(cc).strip().lower():
                return cc
    return None


def main():
    L = ["# quick_check_kedro —— kedro 輸出 報告 capex/opex by year（萬）對 golden", ""]
    L.append(f"   {'家':<8}{'25cap':>9}{'g25cap':>9}{'Δ':>7}{'25opx':>9}{'g25opx':>9}{'Δ':>7}{'24cap':>9}{'23cap':>9}")
    for ent, n in ENTS.items():
        p = ROOT / "data" / ent / "output" / f"company_{n}_kpi_report.parquet"
        if not p.exists():
            L.append(f"   {ent:<8} !! 揾唔到 {p.name}（未跑?）"); continue
        df = pd.read_parquet(p)
        cols = list(df.columns)
        amtc = _pick(cols, ["amount_mop", "amount", "Reported Amount(MOP)", "Val/COArea Crcy", "MOP Amt",
                            "Amount - Amended", "Debit minus Credit", "Entry Voucher Amount/ Expense Amount"])
        coc = _pick(cols, ["final_capex_opex", "capex_opex"])
        rpc = _pick(cols, ["report_period", "year_bucket", "bucket"])
        if not (amtc and coc and rpc):
            L.append(f"   {ent:<8} !! 欄揾唔到 amount={amtc} co={coc} rp={rpc}；cols head={cols[:12]}"); continue
        amt = _num(df[amtc]) / 1e4
        co = df[coc].astype(str).str.strip()
        if ent == "galaxy":      # 正規化：=Capex 先 Capex，其他 Opex
            co = co.where(co.eq("Capex"), "Opex")
        else:
            co = co.map(lambda x: "Capex" if str(x).strip().lower().startswith("cap") else ("Opex" if str(x).strip().lower().startswith("op") else x))
        rp2 = df[rpc].astype(str).str[:2]
        def _s(pref, kind):
            m = rp2.eq(pref) & co.eq(kind)
            return amt[m].sum()
        c25, o25 = _s("25", "Capex"), _s("25", "Opex")
        c24, c23 = _s("24", "Capex"), _s("23", "Capex")
        gc, go = GCAP[ent], GOPX[ent]
        L.append(f"   {ent:<8}{c25:>9,.0f}{gc:>9,.0f}{c25-gc:>7,.0f}{o25:>9,.0f}{go:>9,.0f}{o25-go:>7,.0f}{c24:>9,.0f}{c23:>9,.0f}")
    L.append("")
    L.append("   註：amount = 報告(調整前)。25=report_period prefix 25(含25_24SY/25_23SY)。Δ=我−golden。")
    L.append("   ⚠️ 呢個係 kedro 輸出(未經 prep_tableau 嘅 H 重分類/comp-staff retag)；capex 應該已準，comp/staff 要 prep 後先準。")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
