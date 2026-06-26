r"""wynn_amount_investigation.py — 徹查 wynn 25 調整後 +110 root（read-only，唔郁數）
Run: python scripts\wynn_amount_investigation.py
Out: results\wynn_amount_investigation.txt
逐 stage 追金額：kpi_report parquet 各 amount 欄 by year + 341 問題行特徵。
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT  = ROOT / "results" / "wynn_amount_investigation.txt"
KPI  = ROOT / "data" / "wynn" / "output" / "company_3_kpi_report.parquet"
RAWP = ROOT / "data" / "wynn" / "interim" / "company_3_raw.parquet"


def _num(s): return pd.to_numeric(s, errors="coerce")


def main():
    L = ["# wynn amount 徹查（kpi_report + step0.5 raw parquet）", ""]
    if not KPI.exists():
        L.append(f"!! {KPI} 揾唔到"); OUT.parent.mkdir(exist_ok=True); OUT.write_text("\n".join(L),encoding="utf-8"); print("\n".join(L)); return
    df = pd.read_parquet(KPI)
    L.append(f"kpi_report rows={len(df):,}")

    # 金額相關欄（連 repr 睇空格）
    amt_like = [c for c in df.columns if any(k in str(c).lower() for k in
                ["amount","voucher","expense amount","調整","adjust","original","金額","net"])]
    L += ["", "── 金額相關欄 repr ──"] + [f"   {repr(c)}" for c in amt_like]

    yr = "year" if "year" in df.columns else ("report_period" if "report_period" in df.columns else None)
    bk = "report_period" if "report_period" in df.columns else None
    L.append(f"\nyear 欄={yr}  bucket 欄={bk}")

    # 各 amount 欄 by year（萬）
    if yr:
        L += ["", "── 各金額欄 Σ(萬) by year ──"]
        ycol = df[yr].astype(str).str[:4]
        hdr = "   " + "year".ljust(8) + "".join(c[:18].rjust(20) for c in amt_like)
        L.append(hdr)
        for y in sorted(ycol.unique()):
            m = ycol.eq(y)
            row = "   " + y.ljust(8) + "".join(f"{_num(df.loc[m,c]).sum()/1e4:>20,.0f}" for c in amt_like)
            L.append(row)

    # 341 問題行：amount_mop≈0 但 Entry Voucher≠0（25系列）
    ev = next((c for c in df.columns if "voucher" in str(c).lower() or "expense amount" in str(c).lower()), None)
    if "amount_mop" in df.columns and ev and bk:
        m25 = df[bk].astype(str).str.startswith("25")
        a = _num(df["amount_mop"]).fillna(0); e = _num(df[ev]).fillna(0)
        prob = m25 & (a.abs() <= 0.5) & (e.abs() > 0.5)
        L += ["", f"── 問題行（25系列, amount_mop≈0 但 {ev[:20]}≠0）──",
              f"   行數={int(prob.sum()):,}  Σ{ev[:16]}={e[prob].sum()/1e4:,.0f}萬  Σamount_mop={a[prob].sum()/1e4:,.0f}萬"]
        g = df[prob].copy()
        for fc in ["year","netoff_flag","take_flag2","horizontal_label","row_type","Capex / Opex"]:
            if fc in g.columns:
                vc = g[fc].astype(str).value_counts().head(5)
                L.append(f"   by {fc}: " + " | ".join(f"{k}={v}" for k,v in vc.items()))
        # 呢啲行喺其他金額欄有冇值？
        L.append("   呢批行各金額欄 Σ(萬)：")
        for c in amt_like:
            L.append(f"      {repr(c):<46} {_num(g[c]).sum()/1e4:>12,.0f}")
        # 樣本
        sc = [c for c in ["year", ev, "amount_mop", "original_amount", "adjustment_amount", "調整後金額", "Account", "Expense Description", "netoff_flag"] if c in g.columns]
        L.append("   樣本 6 行：")
        for _, r in g.head(6).iterrows():
            L.append("      " + " | ".join(f"{c[:12]}={str(r[c])[:16]}" for c in sc))

    # step0.5 raw parquet（睇 original_amount 點嚟）
    if RAWP.exists():
        rp = pd.read_parquet(RAWP, columns=None)
        L += ["", "── step0.5 raw parquet ──", f"   rows={len(rp):,}  amount 欄={[c for c in rp.columns if 'amount' in c.lower() or 'voucher' in c.lower() or 'original' in c.lower()][:8]}"]

    OUT.parent.mkdir(exist_ok=True); OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {OUT.relative_to(ROOT)}")

if __name__ == "__main__":
    main()
