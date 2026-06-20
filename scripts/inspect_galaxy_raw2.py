r"""galaxy raw Combine(clean) 嘅項目組標簽 × amount basis 全掃。
group by 基礎|一級+二級標簽（同 人工|二級），sum Reported(調整前)/調整金額/調整後金額。
→ 確認每個 comp/staff component 用 調整前 定 調整後 = golden，同 map 去我哋 H。

Run（Windows）:  python scripts\inspect_galaxy_raw2.py
出 results\inspect_galaxy_raw2.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
F = ROOT / "data" / "galaxy" / "raw" / "project_team" / "galaxy_2025.xlsx"
SHEET = "Combine(clean)"
L: list[str] = [f"# inspect_galaxy_raw2 — {F.name} [{SHEET}] 項目組標簽 × amount（萬）"]


def num(s):
    return pd.to_numeric(s.astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0.0)


def main():
    if not F.exists():
        L.append(f"!! 揾唔到 {F}"); _w(); return
    df = pd.read_excel(F, sheet_name=SHEET, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    rep = num(df.get("Reported Amount(MOP)", 0)) / 1e4   # 萬
    adj = num(df.get("調整金額", 0)) / 1e4
    post = num(df.get("調整後金額", 0)) / 1e4
    df["_rep"] = rep; df["_adj"] = adj; df["_post"] = post
    b1 = df.get("基礎|一級標簽", "").astype(str).str.strip()
    b2 = df.get("基礎|二級標簽", "").astype(str).str.strip()
    df["_b"] = b1 + "|" + b2
    co = df.get("Capex/Opex", "").astype(str).str.strip().str.lower()
    df["_opex"] = ~co.str.contains("capex")

    L.append(f"\n總 Reported(調整前)={rep.sum():,.0f}萬 | 調整後={post.sum():,.0f}萬")

    # ── 基礎|一級 大類 ──
    L.append(f"\n{'='*64}\n## 基礎|一級標簽（大類）opex only：調整前 vs 調整後")
    o = df[df._opex]
    g = o.groupby(b1[df._opex])[["_rep", "_post"]].sum().sort_values("_rep", key=lambda s: s.abs(), ascending=False)
    for k, r in g.iterrows():
        if abs(r["_rep"]) >= 100:
            L.append(f"   {str(k)[:22]:<22} 調整前={r['_rep']:>10,.0f}  調整後={r['_post']:>10,.0f}")

    # ── Comp 細類（基礎|二級）：調整前 vs 調整後（對 golden 客房/會場/食飲/票）──
    comp = o[b1[df._opex].eq("Comp")]
    L.append(f"\n{'='*64}\n## Comp 細類（基礎|二級）opex：調整前 vs 調整後（golden24 客房32,356/會場33,330/食飲5,876/票4,024）")
    gc = comp.groupby(b2[comp.index])[["_rep", "_post"]].sum().sort_values("_rep", key=lambda s: s.abs(), ascending=False)
    for k, r in gc.iterrows():
        L.append(f"   Comp|{str(k)[:18]:<18} 調整前={r['_rep']:>10,.0f}  調整後={r['_post']:>10,.0f}")

    # ── 人工 軸（人工|一級+二級）opex：staff ──
    h1 = df.get("人工|一級標簽", "").astype(str).str.strip()
    h2 = df.get("人工|二級標簽", "").astype(str).str.strip()
    L.append(f"\n{'='*64}\n## 人工|一級標簽 opex：調整前 vs 調整後（golden 25 staff 23,950；24=14,574）")
    gh = o.groupby(h1[df._opex])[["_rep", "_post"]].sum().sort_values("_rep", key=lambda s: s.abs(), ascending=False)
    for k, r in gh.iterrows():
        if abs(r["_rep"]) >= 100:
            L.append(f"   {str(k)[:22]:<22} 調整前={r['_rep']:>10,.0f}  調整後={r['_post']:>10,.0f}")
    # 人工|一級 = staff cost 嘅細類
    sc = o[h1[df._opex].str.contains("staff", case=False, na=False)]
    if len(sc):
        L.append(f"   ── 人工|一級含 staff 嘅二級細分 ──")
        gs = sc.groupby(h2[sc.index])[["_rep", "_post"]].sum().sort_values("_rep", key=lambda s: s.abs(), ascending=False)
        for k, r in gs.head(10).iterrows():
            L.append(f"      staff|{str(k)[:16]:<16} 調整前={r['_rep']:>9,.0f}  調整後={r['_post']:>9,.0f}")
    _w()


def _w():
    out = ROOT / "results" / "inspect_galaxy_raw2.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
