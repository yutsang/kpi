r"""inspect_wynn_adj_ng.py — wynn 逐 bucket × NG：amount_mop / 調整前 / 調整 / 調整後（萬）
Run: python scripts\inspect_wynn_adj_ng.py
In : tableau_combined_25.csv
Out: results\inspect_wynn_adj_ng.txt
判斷 wynn 2025 amount 係調整前定調整後：
  若 amount_mop ≈ golden(報告/post) → amount 已 post，step5 再加 調整 = 重覆 → 改唔加。
  若 amount_mop = golden 之 調整前 → amount 係 pre，調整後 = pre+調整 啱（110萬 = 真淨調整）。
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT  = ROOT / "results" / "inspect_wynn_adj_ng.txt"
CSV  = next((c for c in [ROOT/"tableau_combined_25.csv", Path("tableau_combined_25.csv")] if c.exists()), None)
COLS = ["entity", "year_bucket", "ng_code", "ng_label",
        "amount_mop", "調整前_萬", "調整_萬", "調整後_萬"]


def main():
    if CSV is None:
        OUT.parent.mkdir(exist_ok=True); OUT.write_text("!! csv 揾唔到", encoding="utf-8"); print("no csv"); return
    parts = []
    for ch in pd.read_csv(CSV, usecols=lambda c: c in COLS, chunksize=300_000, dtype=str, encoding="utf-8-sig"):
        ch = ch.rename(columns=lambda c: c.strip())
        ch = ch[ch["entity"] == "wynn"]
        if not len(ch): continue
        ch["amt萬"] = pd.to_numeric(ch["amount_mop"], errors="coerce").fillna(0.0) / 1e4
        for c in ("調整前_萬", "調整_萬", "調整後_萬"):
            ch[c] = pd.to_numeric(ch.get(c), errors="coerce").fillna(0.0)
        ch["ng_label"] = ch["ng_label"].fillna("").astype(str).str.strip()
        ch["ng_code"] = ch["ng_code"].fillna("").astype(str).str.strip()
        ch["year_bucket"] = ch["year_bucket"].fillna("").astype(str).str.strip()
        parts.append(ch.groupby(["year_bucket","ng_code","ng_label"], as_index=False)
                       .agg(amt萬=("amt萬","sum"), 調整前=("調整前_萬","sum"),
                            調整=("調整_萬","sum"), 調整後=("調整後_萬","sum")))
    g = pd.concat(parts, ignore_index=True).groupby(["year_bucket","ng_code","ng_label"], as_index=False).sum()

    L = ["# wynn 逐 bucket × NG（萬MOP）  amount_mop / 調整前 / 調整 / 調整後", ""]
    for bk in ["25", "25_24SY", "25_23SY", "24", "24_23SY", "23"]:
        sub = g[g.year_bucket == bk].sort_values("amt萬", ascending=False)
        if not len(sub): continue
        L += ["", "=" * 78, f"## bucket {bk}", 
              f"   {'NG':<28}{'amount_mop':>12}{'調整前':>12}{'調整':>10}{'調整後':>12}"]
        for _, r in sub.iterrows():
            L.append(f"   {(r.ng_code+' '+r.ng_label)[:28]:<28}{r['amt萬']:>12,.0f}{r['調整前']:>12,.0f}{r['調整']:>10,.0f}{r['調整後']:>12,.0f}")
        L.append(f"   {'— bucket 合計 —':<28}{sub['amt萬'].sum():>12,.0f}{sub['調整前'].sum():>12,.0f}{sub['調整'].sum():>10,.0f}{sub['調整後'].sum():>12,.0f}")

    # 25 系列總計（user 對 golden 嘅 by-NG 通常係呢個範圍）
    s25 = g[g.year_bucket.str.startswith("25")]
    L += ["", "=" * 78, "## 25 系列(25+25_24SY+25_23SY) 合計",
          f"   amount_mop={s25['amt萬'].sum():,.0f}萬  調整前={s25['調整前'].sum():,.0f}  調整={s25['調整'].sum():,.0f}  調整後={s25['調整後'].sum():,.0f}",
          f"   →（調整後 − amount_mop）= {s25['調整後'].sum()-s25['amt萬'].sum():,.0f}萬  ← 若 ≈110萬 = step5 喺25加咗調整金額"]
    OUT.parent.mkdir(exist_ok=True); OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {OUT.relative_to(ROOT)}")

if __name__ == "__main__":
    main()
