r"""inspect_wynn_amount_cols.py — 拆 wynn 金額欄，揾 step5 base(Entry Voucher) vs amount_mop 104萬
Run: python scripts\inspect_wynn_amount_cols.py
Out: results\inspect_wynn_amount_cols.txt
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
        OUT.parent.mkdir(exist_ok=True); OUT.write_text("!! parquet 揾唔到", encoding="utf-8"); print("no pq"); return
    df = pd.read_parquet(pq)
    L.append(f"讀: {pq.name}  rows={len(df):,}  cols={len(df.columns)}")

    # 0) 全欄名（連 repr 睇空格）
    L += ["", "── 0) 全部欄名 (repr) ──"]
    for c in df.columns:
        L.append(f"   {repr(c)}")

    bk = "report_period" if "report_period" in df.columns else ("year_bucket" if "year_bucket" in df.columns else None)
    L += ["", f"bucket 欄 = {bk}"]
    m25 = df[bk].astype(str).str.startswith("25") if bk else pd.Series(True, index=df.index)

    # 1) 金額相關欄 full Σ + bucket25 Σ
    amt_cols = [c for c in df.columns if any(k in str(c).lower() for k in
                ["amount","entry","調整","net","mop","金額","val/co","voucher","expense"])]
    L += ["", "── 1) 金額欄  full Σ(萬)  |  bucket25 Σ(萬) ──"]
    for c in amt_cols:
        s = _num(df[c])
        if s.notna().any():
            L.append(f"   {repr(c):<46} {s.sum()/1e4:>13,.0f} | {_num(df.loc[m25,c]).sum()/1e4:>13,.0f}")
        else:
            L.append(f"   {repr(c):<46} [非數值]")

    # 2) amount_mop vs Entry-Voucher-like 逐行比 (bucket25)
    ev = next((c for c in df.columns if "voucher" in str(c).lower() or "entry" in str(c).lower()), None)
    L += ["", f"── 2) amount_mop vs '{ev}' (bucket25) ──"]
    if "amount_mop" in df.columns and ev:
        d = df[m25]
        a = _num(d["amount_mop"]).fillna(0); e = _num(d[ev]).fillna(0)
        diff = e - a
        L.append(f"   Σamount_mop={a.sum()/1e4:,.0f}萬  Σ{ev[:20]}={e.sum()/1e4:,.0f}萬  Σ差={(e.sum()-a.sum())/1e4:,.0f}萬")
        mm = diff.abs() > 0.005
        L.append(f"   唔等行={int(mm.sum()):,}")
        if mm.any():
            g = d[mm].copy(); g["_d"] = diff[mm]/1e4
            for fc in ["netoff_flag","Net-off","take_flag","take_flag2","final_capex_opex","report_period","horizontal_label"]:
                if fc in g.columns:
                    ser = g.groupby(g[fc].astype(str))["_d"].sum()
                    L.append(f"   by {fc}: " + " | ".join(f"{k}={v:,.0f}萬" for k,v in ser.items()))
            for nc in ["account_desc","description","vendor","project"]:
                if nc in g.columns:
                    ser = g.groupby(g[nc].astype(str))["_d"].sum()
                    ser = ser.reindex(ser.abs().sort_values(ascending=False).index).head(8)
                    L.append(f"   top {nc} (by |差|):")
                    for k,v in ser.items():
                        L.append(f"      {str(k)[:46]:<48} {v:,.0f}萬")
                    break
    else:
        L.append(f"   (amount_mop in df={'amount_mop' in df.columns}; ev={ev})")

    OUT.parent.mkdir(exist_ok=True); OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {OUT.relative_to(ROOT)}")

if __name__ == "__main__":
    main()
