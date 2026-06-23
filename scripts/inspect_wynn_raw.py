r"""inspect_wynn_raw —— 讀 wynn raw excel + 對 CSV，揾 capex-by-year 邊度唔對。

user：wynn raw excel filter capex 分年份 應 209,699/5,611/40，但出嚟錯。
讀 raw 每年（2025=Opex&Capex匯總/2024=sheet0/2023=Sheet1）→ capex/opex × amount，
+ 對 tableau_combined_25.csv 嘅 wynn capex by year_bucket，逐項比。

Run（Windows）:  python scripts\inspect_wynn_raw.py
出 results\inspect_wynn_raw.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "inspect_wynn_raw.txt"

# (year, filename候選, sheet, amount候選, capex候選)
SRC = [
    ("2025", ["wynn_2025.xlsx"], "Opex&Capex匯總",
     ["Entry Voucher Amount/ Expense Amount ", "Entry Voucher Amount/ Expense Amount", "amount"],
     ["Capex / Opex", "Capex/Opex", "capex_opex"]),
    ("2024", ["wynn_2024.xlsx"], 0,
     ["Entry Voucher Amount/ Expense Amount ", "amount", "Amount"],
     ["Capex/Opex", "Capex / Opex", "capex_opex"]),
    ("2023", ["wynn_2023.xlsx"], "Sheet1",
     ["adjusted_amount", "amount", "Amount"],
     ["capex_opex", "Capex/Opex", "Capex / Opex"]),
]


def _find(name):
    for p in [ROOT / "data" / "wynn" / "raw" / name,
              ROOT / "data" / "wynn" / "raw" / name[-4:] / name]:
        if p.exists():
            return p
    hits = list((ROOT / "data" / "wynn").rglob(name))
    return hits[0] if hits else None


def _col(df, cands):
    for c in cands:
        if c in df.columns:
            return c
    for c in df.columns:
        for cand in cands:
            if str(cand).strip().lower() in str(c).strip().lower():
                return c
    return None


def main():
    L = ["# inspect_wynn_raw —— raw excel capex/opex by year vs CSV"]
    # ── raw ──
    for yr, fns, sheet, amts, caps in SRC:
        fp = next((p for fn in fns if (p := _find(fn))), None)
        if not fp:
            L.append(f"\n## {yr}: !! 揾唔到 raw（試 {fns}）"); continue
        try:
            df = pd.read_excel(fp, sheet_name=sheet, dtype=object)
        except Exception as e:
            L.append(f"\n## {yr}: 讀 {fp.name}/{sheet} 失敗: {e}"); continue
        df.columns = [str(c).strip() for c in df.columns]
        ac = _col(df, amts); cc = _col(df, caps)
        L.append(f"\n## {yr}  {fp.name} / {sheet}  rows={len(df):,}")
        L.append(f"   amount欄={ac!r}  capex欄={cc!r}")
        if not ac:
            L.append(f"   欄: {list(df.columns)[:20]}"); continue
        amt = pd.to_numeric(df[ac].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0.0)
        L.append(f"   Σ全部 = {amt.sum():,.0f} ({amt.sum()/1e4:,.1f}萬)")
        if cc:
            g = df.assign(_a=amt).groupby(df[cc].astype(str).str.strip())["_a"].sum().sort_values(key=lambda s: s.abs(), ascending=False)
            for k, v in g.items():
                L.append(f"     {cc}={k[:18]:<18}{v/1e4:>12,.1f}萬")

    # ── CSV wynn capex by year_bucket ──
    if CSV.exists():
        c = pd.read_csv(CSV, dtype=str, keep_default_na=False)
        w = c[c["entity"].astype(str).eq("wynn")].copy()
        w["amt"] = pd.to_numeric(w.get("amount_mop", 0), errors="coerce").fillna(0.0)
        w["cap"] = w.get("final_capex_opex", pd.Series("", index=w.index)).astype(str).str.strip().eq("Capex")
        L.append("\n## CSV wynn capex(final_capex_opex=Capex) by year_bucket（萬）")
        for yb in sorted(w["year_bucket"].astype(str).unique()):
            s = w[w["year_bucket"].astype(str).eq(yb)]
            L.append(f"   {yb:<12}capex={s[s.cap]['amt'].sum()/1e4:>11,.1f}  total={s['amt'].sum()/1e4:>11,.1f}")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
