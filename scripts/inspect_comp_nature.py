r"""inspect_comp_nature —— comp 性質錯位（galaxy 房/票、sjm 會場、mgm 會場）用邊條欄對得返。

HQ comp 拆 5 性質(客房/會場/食品飲料/贈送門票/其他)。我哋 25 性質 vs HQ：
  galaxy 房−2,162/票−2,533 → 餐飲+1,526/其他+1,985；sjm 會場−1,606→客房+2,412；mgm 會場−1,110→食飲+1,775。
查 comp 行（COMP_H）by (comp_type, 項目組H, 現horizontal, account_desc) → 睇邊條欄有正確性質。
Run（Windows）:  python scripts\inspect_comp_nature.py [ent ...]  (default galaxy sjm mgm)
出 results\inspect_comp_nature.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "inspect_comp_nature.txt"
ENTS = sys.argv[1:] or ["galaxy", "sjm", "mgm"]
COMP_H = {"H_HOTEL_ROOM", "H_VENUE", "H_FNB", "H_COMP_TICKET", "H_COMP_OTHER"}
HLAB = {"H_HOTEL_ROOM": "客房", "H_VENUE": "會場", "H_FNB": "食飲", "H_COMP_TICKET": "贈票", "H_COMP_OTHER": "comp其他"}


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# inspect_comp_nature —— comp 性質錯位 source 欄（25,調整前萬）"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["pre"] = pd.to_numeric(df.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    df["y2"] = df["year_bucket"].astype(str).str[:2]
    for c, n in [("entity", "e"), ("horizontal_id", "hid"), ("comp_type", "ct"),
                 ("項目組H", "pth"), ("account_desc", "ad"), ("account_code", "code"), ("is_internal", "ii")]:
        df[n] = _s(df.get(c, pd.Series("", index=df.index)))

    for ent in ENTS:
        e = df[df.e.eq(ent) & df.y2.eq("25") & df.hid.isin(COMP_H)].copy()
        L.append("\n" + "█" * 78)
        L.append(f"█ {ent} 25 comp（COMP_H）Σ={e['pre'].sum():,.0f}萬")
        L.append("█" * 78)
        # 現 H 分佈
        L.append("\n  現 horizontal 分佈：")
        for h, v in e.groupby("hid")["pre"].sum().sort_values(key=lambda s: s.abs(), ascending=False).items():
            L.append(f"     {HLAB.get(h, h):<8}{v:>10,.0f}萬")
        # comp_type 有冇值 + 對唔對得返性質
        L.append("\n  ── by comp_type × 現H（睇 comp_type 有冇正確性質）──")
        g = e.groupby(["ct", "hid"])["pre"].sum().reset_index().sort_values("pre", key=lambda s: s.abs(), ascending=False)
        for _, r in g.head(14).iterrows():
            L.append(f"     ct={(r['ct'] or '(空)')[:22]:<22} H={HLAB.get(r['hid'], r['hid']):<8}{r['pre']:>10,.0f}萬")
        # 項目組H 有冇值
        L.append("\n  ── by 項目組H × 現H（galaxy 用呢個）──")
        g2 = e.groupby(["pth", "hid"])["pre"].sum().reset_index().sort_values("pre", key=lambda s: s.abs(), ascending=False)
        for _, r in g2.head(12).iterrows():
            L.append(f"     pth={(r['pth'] or '(空)')[:24]:<24} H={HLAB.get(r['hid'], r['hid']):<8}{r['pre']:>10,.0f}萬")
        # account_desc top（睇關鍵字定性質）
        L.append("\n  ── top account_desc × 現H ──")
        g3 = e.groupby(["code", "ad", "hid"])["pre"].sum().reset_index().sort_values("pre", key=lambda s: s.abs(), ascending=False)
        for _, r in g3.head(12).iterrows():
            L.append(f"     {(r['code'] or '')[:9]:<9}{(r['ad'] or '(空)')[:30]:<30} H={HLAB.get(r['hid'], r['hid']):<7}{r['pre']:>9,.0f}萬")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
