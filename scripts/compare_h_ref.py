r"""對比我哋 H/V vs 項目組 ref（CSV 入面嘅 項目組H / 項目組V 欄）—— 揾 contradiction。
⚠ 唔照抄（taxonomy 唔同）—— 只標矛盾 / 佢哋分到我哋分唔好嘅位，畀你 review。

直接讀 tableau_combined_25.csv（已有 我哋 horizontal_id + 項目組H），map 項目組H → 我哋 H taxonomy 比對。

Run:  python scripts/compare_h_ref.py
出 results/compare_h_ref.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"

# 項目組H 值（vml 進一步分類 / melco 支出性質 / vml 分類1.1 等）→ 我哋 H_id
HMAP = {
    "工程建設": "H_CONSTRUCTION", "拆除支出": "H_CONSTRUCTION", "建設成本": "H_CONSTRUCTION",
    "設施採購": "H_EQUIP", "器具採購": "H_EQUIP", "一般用品採購": "H_EQUIP",
    "人工成本": "H_LABOR", "人工支出": "H_LABOR", "職工薪酬": "H_LABOR",
    "專業服務費": "H_PROFESSIONAL", "外採服務成本": "H_PROFESSIONAL",
    "食品飲料支出": "H_FNB", "內部附贈資源-食物": "H_FNB", "客房支出": "H_HOTEL_ROOM",
    "內部附贈資源-客房": "H_HOTEL_ROOM", "會場支出": "H_VENUE", "贈票支出": "H_COMP_TICKET",
    "Comp其他": "H_COMP_OTHER", "內部附贈資源-其他": "H_COMP_OTHER",
    "廣告費": "H_ADVERTISING", "媒體費用": "H_ADVERTISING", "營銷費用": "H_ADVERTISING",
    "贊助費用": "H_SPONSORSHIP", "維修與保養費用": "H_MAINTENANCE", "表演者費用": "H_PERFORMER",
    "授權費及版權費": "H_LICENSE",
}


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# compare_h_ref — 我哋 H vs 項目組H（CSV）contradiction（唔照抄，只 review）"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    if "項目組H" not in df.columns:
        L.append("!! csv 冇 項目組H 欄"); _w(L); return
    df["amt"] = pd.to_numeric(df["amount_mop"], errors="coerce").fillna(0.0) / 1e4
    ph = _s(df["項目組H"]); ourh = _s(df["horizontal_id"]); ent = df["entity"].astype(str)
    acc = _s(df["account_desc"]) if "account_desc" in df.columns else pd.Series("", index=df.index)
    for e in sorted(ent.unique()):
        m = ent == e
        cov = (ph[m] != "").mean() * 100
        L.append(f"\n{'='*60}\n── {e}  項目組H 有值 {cov:.0f}% ──")
        if cov < 1:
            L.append("   （冇 項目組H ref，skip）"); continue
        d = df[m].copy()
        d["ph"] = ph[m]; d["exp"] = d["ph"].map(HMAP); d["our"] = ourh[m]; d["acc"] = acc[m]
        mapd = d[(d["exp"].notna()) & (d["ph"] != "")]
        bad = mapd[mapd["exp"] != mapd["our"]]
        L.append(f"   有 map {len(mapd):,} 行；contradiction {len(bad):,} 行 Σ={bad['amt'].abs().sum():,.0f}萬")
        if not len(bad):
            continue
        g = bad.groupby(["ph", "exp", "our"]).agg(s=("amt", "sum"), n=("amt", "size")).reset_index()
        g = g.reindex(g["s"].abs().sort_values(ascending=False).index)
        L.append("   ── 佢哋項目組H（應=我哋H）≠ 我哋實際H，top 12 ──")
        for r in g.head(12).itertuples():
            sa = bad[(bad["ph"] == r.ph) & (bad["our"] == r.our)]["acc"].mode()
            sa = sa.iloc[0] if len(sa) else ""
            L.append(f"     「{str(r.ph)[:10]:<10}」應={r.exp:<14} 我哋={r.our:<14} {r.s:>8,.0f}萬 ({int(r.n)}行) e.g.{str(sa)[:18]}")
    _w(L)


def _w(L):
    out = ROOT / "results" / "compare_h_ref.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
