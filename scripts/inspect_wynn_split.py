r"""inspect_wynn_split —— wynn 2025 raw 點 split 去 25 / 25_24SY / 25_23SY，揾 +107/+26 萬。

user：wynn 25 多 107萬、25_24SY 多 26萬 vs databook。bucket split 喺 kedro（按 raw 嘅 期後/年 欄）。
dump wynn 2025 raw 每個低基數欄（≤25 distinct）× capex sum → 揾邊欄係 split key + 邊個值有差。

Run（Windows）:  python scripts\inspect_wynn_split.py
出 results\inspect_wynn_split.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "inspect_wynn_split.txt"
AMT = ["Entry Voucher Amount/ Expense Amount ", "Entry Voucher Amount/ Expense Amount"]
CAP = ["Capex / Opex", "Capex/Opex"]


def _find(name):
    for p in [ROOT / "data" / "wynn" / "raw" / name]:
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
    L = ["# inspect_wynn_split —— wynn 2025 raw split 欄 × capex"]
    fp = _find("wynn_2025.xlsx")
    if not fp:
        L.append("!! wynn_2025.xlsx 揾唔到"); _w(L); return
    df = pd.read_excel(fp, sheet_name="Opex&Capex匯總", dtype=object)
    df.columns = [str(c).strip() for c in df.columns]
    ac = _col(df, AMT); cc = _col(df, CAP)
    L.append(f"  {fp.name}  rows={len(df):,}  amount={ac!r} capex={cc!r}")
    L.append(f"  全部欄: {list(df.columns)}")
    amt = pd.to_numeric(df[ac].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0.0) if ac else pd.Series(0.0, index=df.index)
    iscap = df[cc].astype(str).str.strip().eq("Capex") if cc else pd.Series(False, index=df.index)
    L.append(f"  Σ全部={amt.sum()/1e4:,.1f}萬  capex={amt[iscap].sum()/1e4:,.1f}萬")

    # 低基數欄（可能係 期後/年/SY split key）
    L.append("\n  ── 低基數欄(≤25 distinct) × capex萬（揾 split key）──")
    for c in df.columns:
        if c in (ac,):
            continue
        vals = df[c].astype(str).str.strip()
        nun = vals.nunique()
        if nun <= 25 and nun > 1:
            g = amt.where(iscap, 0.0).groupby(vals).sum()
            if g.abs().sum() < 1:
                continue
            L.append(f"\n     [{c}] {nun} distinct:")
            for k, v in g.sort_values(key=lambda s: s.abs(), ascending=False).head(12).items():
                L.append(f"        {str(k)[:30]:<30}{v/1e4:>11,.1f}萬")

    # CSV wynn 25-prefix by bucket × capex
    if CSV.exists():
        c = pd.read_csv(CSV, dtype=str, keep_default_na=False)
        w = c[c["entity"].astype(str).eq("wynn") & c["year_bucket"].astype(str).str.startswith("25")].copy()
        w["a"] = pd.to_numeric(w.get("amount_mop", 0), errors="coerce").fillna(0.0)
        w["cap"] = w.get("final_capex_opex", pd.Series("", index=w.index)).astype(str).str.strip().eq("Capex")
        L.append("\n  ── CSV wynn 25-prefix by bucket × capex萬 ──")
        for yb in sorted(w["year_bucket"].astype(str).unique()):
            s = w[w["year_bucket"].astype(str).eq(yb)]
            L.append(f"     {yb:<12}capex={s[s.cap]['a'].sum()/1e4:>11,.1f}  total={s['a'].sum()/1e4:>11,.1f}")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
