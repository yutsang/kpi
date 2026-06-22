r"""inspect_comp_gap_all —— comp 對 user golden（per entity×年），逐個揾 gap 規律（目標 ≤500）。

golden（user 2026-06-22；23/24=調整後, 25=報告/調整前）。basis: 23/24=post, 25=pre。
每 entity×年：我 comp vs golden Δ + by 性質(客房/會場/食飲/贈票/comp其他) + top account（揾 gap 喺邊）。
Run（Windows）:  python scripts\inspect_comp_gap_all.py
出 results\inspect_comp_gap_all.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "inspect_comp_gap_all.txt"
COMP_H = {"H_HOTEL_ROOM", "H_VENUE", "H_FNB", "H_COMP_TICKET", "H_COMP_OTHER"}
HLAB = {"H_HOTEL_ROOM": "客房", "H_VENUE": "會場", "H_FNB": "食飲", "H_COMP_TICKET": "贈票", "H_COMP_OTHER": "comp其他"}
GOLDEN = {  # 23/24=調整後, 25=報告
    "galaxy": {"23": 32204, "24": 55321, "25": 91883}, "wynn": {"23": 14232, "24": 24459, "25": 21488},
    "vml": {"23": 6173, "24": 6335, "25": 9145}, "melco": {"23": 3247, "24": 6374, "25": 8424},
    "mgm": {"23": 10127, "24": 9073, "25": 7914}, "sjm": {"23": 2970, "24": 6092, "25": 3575},
}
ENTS = ["galaxy", "wynn", "vml", "melco", "mgm", "sjm"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# inspect_comp_gap_all —— comp vs golden（萬；23/24調整後,25報告）目標≤500"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["pre"] = pd.to_numeric(df.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    df["post"] = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    df["y2"] = df["year_bucket"].astype(str).str[:2]
    df["e"] = df["entity"].astype(str)
    df["hid"] = _s(df.get("horizontal_id", pd.Series("", index=df.index)))
    df["ad"] = _s(df.get("account_desc", pd.Series("", index=df.index)))
    df["code"] = _s(df.get("account_code", pd.Series("", index=df.index)))
    comp = df[df["hid"].isin(COMP_H)].copy()

    # 總覽表
    L.append("\n## 總覽 Δ=我−golden（>±500 標 ◄）")
    L.append(f"   {'entity':<8}{'23我':>8}{'g23':>8}{'Δ23':>7}{'24我':>8}{'g24':>8}{'Δ24':>7}{'25我':>8}{'g25':>8}{'Δ25':>7}")
    for ent in ENTS:
        cells, flags = "", ""
        for yr in ("23", "24", "25"):
            sub = comp[comp.e.eq(ent) & comp.y2.eq(yr)]
            mine = sub["pre"].sum() if yr == "25" else sub["post"].sum()
            g = GOLDEN[ent][yr]; d = mine - g
            mk = "◄" if abs(d) > 500 else " "
            cells += f"{mine:>8,.0f}{g:>8,.0f}{d:>6,.0f}{mk}"
        L.append(f"   {ent:<8}{cells}")

    # 逐家×年（只列 |Δ|>500）by 性質 + top account
    for ent in ENTS:
        for yr in ("23", "24", "25"):
            sub = comp[comp.e.eq(ent) & comp.y2.eq(yr)]
            basis = "pre" if yr == "25" else "post"
            mine = sub[basis].sum(); g = GOLDEN[ent][yr]; d = mine - g
            if abs(d) <= 500:
                continue
            L.append(f"\n{'='*64}\n## {ent} {yr}（{basis}）我={mine:,.0f} golden={g:,.0f} Δ={d:+,.0f} ◄")
            L.append("   by 性質：")
            for h, v in sub.groupby("hid")[basis].sum().sort_values(key=lambda s: s.abs(), ascending=False).items():
                L.append(f"     {HLAB.get(h, h):<9}{v:>10,.0f}萬")
            L.append("   top account（睇 gap 喺邊）：")
            for (c, a), v in sub.groupby(["code", "ad"])[basis].sum().sort_values(key=lambda s: s.abs(), ascending=False).head(8).items():
                L.append(f"     {(c or '')[:10]:<10}{(a or '(空)')[:32]:<32}{v:>9,.0f}萬")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
