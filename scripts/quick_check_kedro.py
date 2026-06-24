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
# 每家 kedro 輸出嘅 amount 欄名（= conf amount value；wynn 有尾空格）
AMT = {"galaxy": "amount_mop", "sjm": "Val/COArea Crcy", "wynn": "Entry Voucher Amount/ Expense Amount ",
       "vml": "MOP Amt", "melco": "Amount - Amended", "mgm": "Debit minus Credit"}
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
        amtc = AMT[ent] if AMT[ent] in cols else _pick(cols, [AMT[ent].strip(), "amount_mop", "amount"])
        coc = _pick(cols, ["final_capex_opex", "capex_opex"])
        rpc = _pick(cols, ["report_period", "year_bucket", "bucket"])
        if not (amtc and coc and rpc):
            L.append(f"   {ent:<8} !! 欄揾唔到 amount={amtc} co={coc} rp={rpc}；cols={cols[:14]}"); continue
        amt = _num(df[amtc]) / 1e4
        rp2 = df[rpc].astype(str).str[:2]

        def _capx(co_series, normalize_galaxy=False):
            c = co_series.astype(str).str.strip()
            if normalize_galaxy:
                c = c.where(c.eq("Capex"), "Opex")
            else:
                c = c.map(lambda x: "Capex" if str(x).strip().lower().startswith("cap") else "Opex")
            return c

        co = _capx(df[coc], normalize_galaxy=(ent == "galaxy"))
        c25 = amt[rp2.eq("25") & co.eq("Capex")].sum(); o25 = amt[rp2.eq("25") & co.eq("Opex")].sum()
        c24 = amt[rp2.eq("24") & co.eq("Capex")].sum(); c23 = amt[rp2.eq("23") & co.eq("Capex")].sum()
        gc, go = GCAP[ent], GOPX[ent]
        L.append(f"   {ent:<8}{c25:>9,.0f}{gc:>9,.0f}{c25-gc:>7,.0f}{o25:>9,.0f}{go:>9,.0f}{o25-go:>7,.0f}{c24:>9,.0f}{c23:>9,.0f}  (amt={amtc[:18]},Σ={amt.sum():,.0f})")
        # galaxy 另用 Capex/Opex 欄對照（睇邊個 = golden）
        if ent == "galaxy":
            cocol = _pick(cols, ["Capex/Opex", "Capex / Opex"])
            if cocol:
                co2 = _capx(df[cocol])
                c25b = amt[rp2.eq("25") & co2.eq("Capex")].sum()
                L.append(f"            ↳ 用「{cocol}」欄: 25cap={c25b:,.0f}（對 golden {gc:,.0f}）")
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
