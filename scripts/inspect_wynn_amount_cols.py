r"""inspect_wynn_amount_cols.py — 拆 wynn 金額欄，揾 step5 base(Entry Voucher) vs amount_mop 嗰 104萬
Run: python scripts\inspect_wynn_amount_cols.py
Out: results\inspect_wynn_amount_cols.txt
讀 wynn kpi_report parquet（row-level），bucket 25 比較各金額欄 + 揾 amount_mop≠EntryVoucher 嘅行。
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT  = ROOT / "results" / "inspect_wynn_amount_cols.txt"
PQ_CANDS = [ROOT/"data"/"wynn"/"output"/"company_3_kpi_report.parquet",
            ROOT/"data"/"wynn"/"interim"/"company_3_tagged_rows.parquet"]


def _num(s): return pd.to_numeric(s, errors="coerce")


def main():
    pq = next((p for p in PQ_CANDS if p.exists()), None)
    L = ["# wynn 金額欄拆解（揾 104萬 gap）", ""]
    if pq is None:
        L.append("!! kpi_report/tagged_rows parquet 揾唔到"); OUT.parent.mkdir(exist_ok=True); OUT.write_text("\n".join(L),encoding="utf-8"); print("\n".join(L)); return
    df = pd.read_parquet(pq)
    L.append(f"讀: {pq.name}  rows={len(df):,}")

    # bucket 欄
    bk = "report_period" if "report_period" in df.columns else ("year_bucket" if "year_bucket" in df.columns else None)
    L.append(f"bucket 欄: {bk}")

    # 所有金額相關欄
    amt_cols = [c for c in df.columns if any(k in str(c).lower() for k in
                ["amount","entry","調整","net","mop","金額","val/co"]) ]
    L += ["", "── 金額相關欄（全 df Σ） ──"]
    for c in amt_cols:
        s = _num(df[c])
        if s.notna().any():
            L.append(f"   {repr(c):<48} Σ={s.sum()/1e4:>14,.0f}萬  非空={int(s.notna().sum()):,}")
        else:
            L.append(f"   {repr(c):<48} [非數值]")

    # bucket 25 系列
    if bk:
        m25 = df[bk].astype(str).str.startswith("25")
        d = df[m25]
        L += ["", f"── bucket 25 系列（{int(m25.sum()):,} 行）各金額欄 Σ ──"]
        for c in amt_cols:
            s = _num(d[c])
            if s.notna().any():
                L.append(f"   {repr(c):<48} Σ={s.sum()/1e4:>14,.0f}萬")

        # amount_mop vs Entry Voucher 逐行比
        ev = next((c for c in df.columns if "entry voucher" in str(c).lower()), None)
        if "amount_mop" in df.columns and ev:
            a = _num(d["amount_mop"]).fillna(0); e = _num(d[ev]).fillna(0)
            diff = (e - a)
            mm = diff.abs() > 0.005
            L += ["", f"── amount_mop vs '{ev}'（bucket25）──",
                  f"   Σ amount_mop={a.sum()/1e4:,.0f}萬   Σ EntryVoucher={e.sum()/1e4:,.0f}萬   Σ差={(e.sum()-a.sum())/1e4:,.0f}萬",
                  f"   唔等嘅行={int(mm.sum()):,}"]
            if mm.any():
                cols_show = [c for c in ["account_code","account_desc","description","Net-off","netoff_flag","take_flag","horizontal_label","vertical_label"] if c in d.columns]
                g = d[mm].copy(); g["_diff萬"] = diff[mm]/1e4
                # by netoff / take_flag 滙總
                for fc in ["Net-off","netoff_flag","take_flag"]:
                    if fc in g.columns:
                        L.append(f"   差額 by {fc}: " + " | ".join(f"{k}={v/1e4:,.0f}萬" for k,v in g.groupby(g[fc].astype(str))["_diff萬"].sum().items()))
                L.append("   差額 top account_desc:")
                for k,v in g.groupby(g.get("account_desc","?").astype(str))["_diff萬"].sum().abs().sort_values(ascending=False).head(10).items():
                    L.append(f"      {str(k)[:40]:<42} {v:,.0f}萬")
                L.append("   樣本 5 行:")
                for _, r in g.head(5).iterrows():
                    L.append("      " + " | ".join(f"{c}={str(r[c])[:22]}" for c in cols_show) + f" | diff={r['_diff萬']:,.0f}萬")

    OUT.parent.mkdir(exist_ok=True); OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {OUT.relative_to(ROOT)}")

if __name__ == "__main__":
    main()
