r"""inspect_sjm_promo —— sjm 推廣費(72040050)by description/科目明細，揾「邊類→邊個comp性質」規則。

user：sjm comp = 推廣費 subset，用 desc 揾規則對齊 golden（23:房752/會場678/食飲1,346/票0/其他194）。
dump 推廣費 rows by description + 現H + amount，揾 酒店/客房、餐飲/F&B、禮品/禮劵/票、活動/場地 等關鍵字。

Run（Windows）:  python scripts\inspect_sjm_promo.py
出 results\inspect_sjm_promo.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "inspect_sjm_promo.txt"
PROMO = r"72040050|推廣費|公司推廣|Promotion Expense|推廣活動"
# 性質 keyword 試探
KWS = {
    "客房": r"酒店|客房|住宿|hotel|room|房",
    "食飲": r"餐|食|飲|宴|f&b|food|beverage|dinner|菜",
    "贈票": r"禮品|禮劵|礼品|礼券|票|ticket|gift|voucher|入場",
    "會場": r"活動|場地|場館|會場|展|venue|event|exhibition|路演|展銷",
    "其他": r"",
}


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# inspect_sjm_promo —— sjm 推廣費 by desc 揾性質規則"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    e = df[df["entity"].astype(str).eq("sjm")].copy()
    e["post"] = pd.to_numeric(e.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    e["pre"] = pd.to_numeric(e.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    e["y2"] = e["year_bucket"].astype(str).str[:2]
    for c, n in [("horizontal_id", "hid"), ("account_code", "code"), ("account_desc", "ad"),
                 ("description", "ds"), ("subproject", "sp"), ("科目明細", "km"),
                 ("project", "pj"), ("ng_label", "ng")]:
        e[n] = _s(e.get(c, pd.Series("", index=e.index)))
    e["blob"] = (e["ad"] + " " + e["ds"] + " " + e["sp"] + " " + e["km"] + " " + e["pj"])
    promo = e[e["code"].str.contains("72040050", na=False) | e["ad"].str.contains(PROMO, case=False, regex=True, na=False)]

    for yr in ("23", "24", "25"):
        s = promo[promo.y2.eq(yr)]
        b = "pre" if yr == "25" else "post"
        L.append(f"\n{'='*64}\n## sjm 推廣費 {yr}（{b}）Σ={s[b].sum():,.0f}萬  by 現H：")
        for h, v in s.groupby("hid")[b].sum().sort_values(key=lambda x: x.abs(), ascending=False).items():
            L.append(f"     {h:<16}{v:>9,.0f}萬")
        # 試性質 keyword 命中
        L.append(f"   ── desc 性質 keyword 命中（睇邊類夠分性質）──")
        _used = pd.Series(False, index=s.index)
        for nat, kw in KWS.items():
            if not kw:
                m = ~_used
            else:
                m = s["blob"].str.contains(kw, case=False, regex=True, na=False) & ~_used
            _used = _used | m
            L.append(f"     {nat:<6}命中 {int(m.sum()):>5}行 Σ={s.loc[m, b].sum():>8,.0f}萬")
        # top description 樣本
        L.append(f"   ── top description 樣本（睇有冇性質字眼）──")
        for d, v in s.groupby("ds")[b].sum().sort_values(key=lambda x: x.abs(), ascending=False).head(10).items():
            L.append(f"     {(d or '(空)')[:46]:<46}{v:>8,.0f}萬")
        # by NG（fit golden 另一路）
        L.append(f"   ── by NG（fit golden NG split 用）──")
        for ng, v in s.groupby("ng")[b].sum().sort_values(key=lambda x: x.abs(), ascending=False).head(8).items():
            incomp = s[(s.ng.eq(ng)) & s.hid.isin({"H_HOTEL_ROOM", "H_VENUE", "H_FNB", "H_COMP_TICKET", "H_COMP_OTHER"})][b].sum()
            L.append(f"     {(ng or '(空)')[:14]:<14}Σ={v:>8,.0f}  其中已comp={incomp:,.0f}")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
