r"""對比我哋 H vs 項目組自己嘅 H-ref（vml 進一步分類 / melco 支出性質），揾 contradiction。
⚠ 唔照抄（taxonomy 唔同）—— 只標「佢哋分到、我哋分唔好」或矛盾嘅位，畀你 review。

讀 raw（佢哋 H-ref + unique_id + 金額）join 我哋 parquet（horizontal_id），
map 佢哋 H-ref → 我哋 H taxonomy，比對，出 contradiction by 金額。

Run:  python scripts/compare_h_ref.py
出 results/compare_h_ref.txt  ← paste 返嚜
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd, yaml

ROOT = Path(__file__).resolve().parent.parent

# entity: (company_n, raw_file, sheet, H-ref col, raw uid col, amount col, {他們值: 我哋H_id})
CFG = {
    "vml": (4, "vml_2025.xlsx", "25JE", "進一步分類", "Ref number", "MOP Amt", {
        "工程建設": "H_CONSTRUCTION", "拆除支出": "H_CONSTRUCTION", "設施採購": "H_EQUIP",
        "器具採購": "H_EQUIP", "人工成本": "H_LABOR", "專業服務費": "H_PROFESSIONAL",
        "食品飲料支出": "H_FNB", "客房支出": "H_HOTEL_ROOM", "會場支出": "H_VENUE",
        "贈票支出": "H_COMP_TICKET", "Comp其他": "H_COMP_OTHER", "廣告費": "H_ADVERTISING",
        "媒體費用": "H_ADVERTISING", "贊助費用": "H_SPONSORSHIP"}),
    "melco": (5, "melco_2025.xlsx", "Data", "支出性質", "KPMG唯一識別碼", "Amount - Amended", {
        "職工薪酬": "H_LABOR", "表演者費用": "H_PERFORMER", "維修與保養費用": "H_MAINTENANCE",
        "營銷費用": "H_ADVERTISING", "一般用品採購": "H_EQUIP", "差旅費": "H_LABOR",
        "贊助費用": "H_SPONSORSHIP", "授權費及版權費": "H_LICENSE", "建設成本": "H_CONSTRUCTION"}),
}


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# compare_h_ref — 我哋 H vs 項目組 H-ref（vml 進一步分類 / melco 支出性質）contradiction"]
    for ent, (n, rf, sh, hcol, uidc, amtc, mp) in CFG.items():
        L.append(f"\n{'='*66}\n── {ent} （H-ref = {hcol}）──")
        rp = ROOT / "data" / ent / "raw" / rf
        pp = ROOT / "data" / ent / "output" / f"company_{n}_kpi_report.parquet"
        if not rp.exists() or not pp.exists():
            L.append("   !! raw / parquet 揾唔到"); continue
        raw = pd.read_excel(rp, sheet_name=sh, dtype=str)
        if hcol not in raw.columns or uidc not in raw.columns:
            L.append(f"   !! raw 冇 {hcol}/{uidc}"); continue
        raw = raw[[uidc, hcol]].copy()
        raw["uid"] = _s(raw[uidc]); raw["href"] = _s(raw[hcol])
        pq = pd.read_parquet(pp)
        uidp = "unique_id" if "unique_id" in pq.columns else None
        if not uidp or "horizontal_id" not in pq.columns:
            L.append("   !! parquet 冇 unique_id/horizontal_id"); continue
        amt = pd.to_numeric(pq[amtc], errors="coerce").fillna(0.0) / 1e4 if amtc in pq.columns else \
            pd.to_numeric(pq.get("amount_mop", 0), errors="coerce").fillna(0.0) / 1e4
        pq = pd.DataFrame({"uid": _s(pq[uidp]), "ourH": _s(pq["horizontal_id"]), "a": amt})
        m = pq.merge(raw[["uid", "href"]], on="uid", how="inner")
        L.append(f"   join: {len(m):,}/{len(pq):,} parquet 行對到 raw")
        if not len(m):
            continue
        m["expH"] = m["href"].map(mp)
        cmp = m[m["expH"].notna()]
        bad = cmp[cmp["expH"] != cmp["ourH"]]
        L.append(f"   有 map 嘅 {len(cmp):,} 行，contradiction {len(bad):,} 行 Σ={bad['a'].abs().sum():,.0f}萬")
        g = bad.groupby(["href", "expH", "ourH"])["a"].agg(["sum", "size"]).reset_index()
        g = g.reindex(g["sum"].abs().sort_values(ascending=False).index)
        L.append("   ── 佢哋H-ref（應=我哋H）≠ 我哋實際H，by 金額 top 15 ──")
        for r in g.head(15).itertuples():
            L.append(f"     「{str(r.href)[:12]:<12}」應={r.expH:<14} 但我哋={r.ourH:<14} {r.sum/1:>8,.0f}萬 ({int(r.size)}行)")
    _w(L)


def _w(L):
    out = ROOT / "results" / "compare_h_ref.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
