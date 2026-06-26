r"""inspect_wynn_rowdiff.py — wynn 25系列 逐行 amount_mop vs 調整前_萬 差，揾嗰 129萬
Run: python scripts\inspect_wynn_rowdiff.py
Out: results\inspect_wynn_rowdiff.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT  = ROOT / "results" / "inspect_wynn_rowdiff.txt"
CSV  = next((c for c in [ROOT/"tableau_combined_25.csv", Path("tableau_combined_25.csv")] if c.exists()), None)
COLS = ["entity","year_bucket","amount_mop","調整前_萬","調整_萬","調整後_萬",
        "account_desc","description","vendor","horizontal_label","ng_label","row_type"]


def main():
    if CSV is None:
        OUT.parent.mkdir(exist_ok=True); OUT.write_text("!! csv 揾唔到", encoding="utf-8"); print("no csv"); return
    parts = []
    for ch in pd.read_csv(CSV, usecols=lambda c: c in COLS, chunksize=300_000, dtype=str, encoding="utf-8-sig"):
        ch = ch.rename(columns=lambda c: c.strip())
        ch = ch[(ch["entity"]=="wynn") & ch["year_bucket"].astype(str).str.startswith("25")]
        if len(ch): parts.append(ch)
    d = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame(columns=COLS)
    d["amt萬"] = pd.to_numeric(d["amount_mop"], errors="coerce").fillna(0)/1e4
    d["pre"]   = pd.to_numeric(d["調整前_萬"], errors="coerce").fillna(0)
    d["diff"]  = d["pre"] - d["amt萬"]

    L = ["# wynn 25系列 逐行 調整前_萬 − amount_mop_萬", ""]
    L.append(f"行數={len(d):,}  Σamount_mop={d['amt萬'].sum():,.0f}萬  Σ調整前={d['pre'].sum():,.0f}萬  Σ差={d['diff'].sum():,.0f}萬")
    mm = d["diff"].abs() > 0.005
    L.append(f"差≠0 行={int(mm.sum()):,}")
    g = d[mm].copy()
    if len(g):
        # 拆解：amount_mop=0 但 調整前≠0 嘅行（最可能）vs 兩者都≠0 但唔等
        z = g[g["amt萬"].abs()<=0.005]
        nz = g[g["amt萬"].abs()>0.005]
        L.append(f"  其中 amount_mop≈0 但 調整前≠0: {len(z):,} 行  Σ調整前={z['pre'].sum():,.0f}萬  Σ差={z['diff'].sum():,.0f}萬")
        L.append(f"  其中 兩者皆≠0 但唔等:        {len(nz):,} 行  Σ差={nz['diff'].sum():,.0f}萬")
        for fc in ["row_type","horizontal_label","ng_label","year_bucket"]:
            if fc in g.columns:
                s = g.groupby(g[fc].astype(str))["diff"].sum().sort_values(key=abs, ascending=False).head(8)
                L.append(f"  差 by {fc}: " + " | ".join(f"{k}={v:,.0f}萬" for k,v in s.items()))
        for nc in ["account_desc","description"]:
            if nc in g.columns:
                s = g.groupby(g[nc].astype(str))["diff"].sum().sort_values(key=abs, ascending=False).head(10)
                L.append(f"  差 top {nc}:")
                for k,v in s.items(): L.append(f"     {str(k)[:50]:<52} {v:,.0f}萬")
                break
        L.append("  樣本 8 行 (diff 最大):")
        for _, r in g.reindex(g["diff"].abs().sort_values(ascending=False).index).head(8).iterrows():
            L.append(f"     pre={r['pre']:,.0f} amt={r['amt萬']:,.0f} diff={r['diff']:,.0f} | {str(r.get('account_desc',''))[:24]} | {str(r.get('description',''))[:30]} | {r.get('row_type','')}")
    OUT.parent.mkdir(exist_ok=True); OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {OUT.relative_to(ROOT)}")

if __name__ == "__main__":
    main()
