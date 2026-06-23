r"""inspect_wynn_sy —— wynn SY bucket 唔 tie 診斷。

wynn year_split = direct_filter：每行直接抄 raw excel 一條欄（是否2025年/是否2024年/year）。
bucket 靠 EXACT string 比對 YAML buckets，對唔到就跌 default（25/24）。
→ dump 嗰幾條 filter 欄嘅每個 distinct 值 × 金額，標明 → 邊個 bucket（或 ⚠ 跌 default），
   再對 CSV wynn by year_bucket，睇邊度差。

Run（Windows）:  python scripts\inspect_wynn_sy.py
出 results\inspect_wynn_sy.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "inspect_wynn_sy.txt"

# (year, file, sheet, amount候選, filter欄, {欄值: bucket}, default)
SRC = [
    ("2025", "wynn_2025.xlsx", "Opex&Capex匯總",
     ["Entry Voucher Amount/ Expense Amount ", "Entry Voucher Amount/ Expense Amount", "amount"],
     "是否2025年", {"2025年度": "25", "2024年度": "25_24SY", "2023年度": "25_23SY"}, "25"),
    ("2024", "wynn_2024.xlsx", 0,
     ["Entry Voucher Amount/ Expense Amount ", "amount", "Amount"],
     "是否2024年", {"2024年度": "24", "2023年計劃2024年完成情況": "24_23SY"}, "24"),
    ("2023", "wynn_2023.xlsx", "Sheet1",
     ["adjusted_amount", "amount", "Amount"],
     "year", {"2023": "23", "2023.0": "23"}, "23"),
]


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


def _fcol(df, requested):
    """同 kedro _resolve_filter_col：exact 優先，否則 unique substring。"""
    if requested in df.columns:
        return requested
    cands = [c for c in df.columns if requested in str(c)]
    return cands[0] if len(cands) == 1 else (cands[0] if cands else None)


def main():
    L = ["# inspect_wynn_sy —— wynn SY bucket 診斷（raw filter欄 × 金額 → bucket）"]
    for yr, fn, sheet, amts, fcol_req, buckets, default in SRC:
        fp = _find(fn)
        if not fp:
            L.append(f"\n## {yr}: !! 揾唔到 {fn}"); continue
        try:
            df = pd.read_excel(fp, sheet_name=sheet, dtype=object)
        except Exception as e:
            L.append(f"\n## {yr}: 讀 {fn}/{sheet} 失敗: {e}"); continue
        df.columns = [str(c).strip() for c in df.columns]
        ac = _col(df, amts)
        fcol = _fcol(df, fcol_req)
        L.append(f"\n{'='*70}\n## {yr}  {fp.name}  rows={len(df):,}")
        L.append(f"   amount欄={ac!r}  filter欄(req={fcol_req!r})={fcol!r}  default={default}")
        if not fcol:
            L.append(f"   !! filter 欄揾唔到！全部行會跌 default={default}。欄: {list(df.columns)[:25]}")
            continue
        amt = pd.to_numeric(df[ac].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0.0) if ac else pd.Series(0.0, index=df.index)
        # 完全照 kedro：astype(str).str.strip()，對 buckets 嘅 str(v).strip()
        fval = df[fcol].astype(str).str.strip()
        exp = {str(k).strip(): v for k, v in buckets.items()}
        L.append(f"   Σ全部={amt.sum()/1e4:,.1f}萬")
        L.append(f"   ── {fcol} 每個 distinct 值 × 金額（標 → bucket / ⚠跌default）──")
        g = amt.groupby(fval).sum().sort_values(key=lambda s: s.abs(), ascending=False)
        nrows = df[fcol].astype(str).str.strip().value_counts()
        bk_tot = {}
        for k, v in g.items():
            bk = exp.get(k, f"⚠default={default}")
            bk_clean = exp.get(k, default)
            bk_tot[bk_clean] = bk_tot.get(bk_clean, 0.0) + v
            flag = "" if k in exp else "   ⚠ 跌 default（exact 對唔到 YAML）"
            kk = repr(k)[:46] if k.strip() else "(空白)"
            L.append(f"      {kk:<48}{v/1e4:>11,.1f}萬  {nrows.get(k,0):>7,}行 → {bk}{flag}")
        L.append(f"   ── 滙總到 bucket ──")
        for bk, v in sorted(bk_tot.items()):
            L.append(f"      {bk:<12}{v/1e4:>12,.1f}萬")

    # CSV wynn by year_bucket
    if CSV.exists():
        c = pd.read_csv(CSV, dtype=str, keep_default_na=False)
        w = c[c["entity"].astype(str).eq("wynn")].copy()
        w["amt"] = pd.to_numeric(w.get("amount_mop", 0), errors="coerce").fillna(0.0)
        w["pre"] = pd.to_numeric(w.get("調整前_萬", 0), errors="coerce").fillna(0.0)
        w["cap"] = w.get("final_capex_opex", pd.Series("", index=w.index)).astype(str).str.strip().eq("Capex")
        L.append(f"\n{'='*70}\n## CSV wynn by year_bucket（萬）")
        L.append(f"   {'bucket':<12}{'amount_mop':>14}{'調整前':>12}{'capex(前)':>12}")
        for yb in sorted(w["year_bucket"].astype(str).unique()):
            s = w[w["year_bucket"].astype(str).eq(yb)]
            L.append(f"   {yb:<12}{s['amt'].sum()/1e4:>14,.1f}{s['pre'].sum():>12,.1f}{s[s.cap]['pre'].sum():>12,.1f}")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
