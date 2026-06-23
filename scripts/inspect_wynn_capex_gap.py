r"""inspect_wynn_capex_gap —— wynn 25-prefix capex：Tableau 209806 vs golden 209702 (+104) 揾源頭。

四個 measure 逐 bucket（25 / 25_24SY / 25_23SY），capex only：
  amount_mop/1e4、調整前_萬、調整_萬、調整後_萬、n_rows
→ 即刻見邊個 measure = 209702（golden=調整前/報告），邊個 = 209806（你 Tableau），+104 = ?
+ raw 2025 excel capex by 是否2025年 值（對返 golden）。

Run（Windows）:  python scripts\inspect_wynn_capex_gap.py
出 results\inspect_wynn_capex_gap.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "inspect_wynn_capex_gap.txt"
AMT25 = ["Entry Voucher Amount/ Expense Amount ", "Entry Voucher Amount/ Expense Amount", "amount"]
CAP25 = ["Capex / Opex", "Capex/Opex", "capex_opex"]


def _find(name):
    p = ROOT / "data" / "wynn" / "raw" / name
    if p.exists(): return p
    hits = list((ROOT / "data" / "wynn").rglob(name))
    return hits[0] if hits else None


def _col(df, cands):
    for c in cands:
        if c in df.columns: return c
    for c in df.columns:
        for cand in cands:
            if str(cand).strip().lower() in str(c).strip().lower(): return c
    return None


def _num(s):
    return pd.to_numeric(pd.Series(s).astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0.0)


def main():
    L = ["# inspect_wynn_capex_gap —— wynn 25-prefix capex 四measure × bucket（揾 +104）"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    c = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    w = c[c["entity"].astype(str).eq("wynn")].copy()
    w["amt"] = pd.to_numeric(w.get("amount_mop", 0), errors="coerce").fillna(0.0) / 1e4
    w["pre"] = pd.to_numeric(w.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    w["adj"] = pd.to_numeric(w.get("調整_萬", 0), errors="coerce").fillna(0.0)
    w["post"] = pd.to_numeric(w.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    w["report"] = pd.to_numeric(w.get("對數金額_萬", 0), errors="coerce").fillna(0.0)
    w["cap"] = w.get("final_capex_opex", pd.Series("", index=w.index)).astype(str).str.strip().eq("Capex")
    yb = w["year_bucket"].astype(str)

    for scope, mask in [("CAPEX only", w["cap"]), ("ALL(capex+opex)", pd.Series(True, index=w.index))]:
        L.append(f"\n{'='*78}\n## CSV wynn {scope} by bucket（萬）")
        L.append(f"   {'bucket':<12}{'amount_mop/1e4':>15}{'調整前':>12}{'調整':>10}{'調整後':>12}{'對數金額':>12}{'n':>9}")
        sub = w[mask]
        for b in ["25", "25_24SY", "25_23SY", "24", "24_23SY", "23"]:
            s = sub[yb.eq(b)]
            if len(s) == 0: continue
            L.append(f"   {b:<12}{s['amt'].sum():>15,.1f}{s['pre'].sum():>12,.1f}{s['adj'].sum():>10,.1f}"
                     f"{s['post'].sum():>12,.1f}{s['report'].sum():>12,.1f}{len(s):>9,}")
        # 25-prefix 小計
        s25 = sub[yb.str.startswith("25")]
        L.append(f"   {'25合計':<12}{s25['amt'].sum():>15,.1f}{s25['pre'].sum():>12,.1f}{s25['adj'].sum():>10,.1f}"
                 f"{s25['post'].sum():>12,.1f}{s25['report'].sum():>12,.1f}{len(s25):>9,}")

    L.append(f"\n   ⭐ 對照你提供：bucket1 Tableau=209,806  golden=209,702（Δ+104）；bucket2 5,637 vs 5,611（Δ+26）")
    L.append(f"   → 上表邊一 column 嘅 25 = 209,702 就係 golden basis；邊個 = 209,806 就係你而家 Tableau 用緊嘅。")

    # raw 2025 excel：capex by 是否2025年
    fp = _find("wynn_2025.xlsx")
    if fp:
        try:
            df = pd.read_excel(fp, sheet_name="Opex&Capex匯總", dtype=object)
            df.columns = [str(x).strip() for x in df.columns]
            ac = _col(df, AMT25); cc = _col(df, CAP25)
            fc = next((x for x in df.columns if "是否2025" in str(x)), None)
            L.append(f"\n{'='*78}\n## raw wynn_2025.xlsx  amount={ac!r} capex={cc!r} 是否2025年={fc!r}")
            if ac and cc and fc:
                amt = _num(df[ac]); iscap = df[cc].astype(str).str.strip().eq("Capex")
                fv = df[fc].astype(str).str.strip()
                g = amt.where(iscap, 0.0).groupby(fv).sum().sort_values(key=lambda s: s.abs(), ascending=False)
                L.append(f"   ── capex by 是否2025年（raw amount /1e4 = 萬）──")
                for k, v in g.items():
                    L.append(f"      {(repr(k)[:40]):<42}{v/1e4:>14,.1f}萬")
                L.append(f"   capex 總 = {amt[iscap].sum()/1e4:,.1f}萬")
        except Exception as e:
            L.append(f"\n## raw 讀取失敗: {e}")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
