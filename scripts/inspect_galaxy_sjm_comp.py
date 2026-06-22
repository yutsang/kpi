r"""inspect_galaxy_sjm_comp —— galaxy(comp under)+ sjm(推廣費分佈) 嘅 comp gap 源。

galaxy：comp 似嘢但唔喺 COMP_H（漏咗 → under）by year。
sjm：推廣費(公司推廣活動/Promotion)across 全 H by year —— golden 當 comp，但 23 太少(−1,714)，
     可能落咗廣告，要撈返。
Run（Windows）:  python scripts\inspect_galaxy_sjm_comp.py
出 results\inspect_galaxy_sjm_comp.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "inspect_galaxy_sjm_comp.txt"
COMP_H = {"H_HOTEL_ROOM", "H_VENUE", "H_FNB", "H_COMP_TICKET", "H_COMP_OTHER"}
COMP_KW = r"Comp|Lodging|Venue License|Comp Food|Comp Beverage|贈|Ticket|Rental Hall|Meeting Room|Service Charge|Special Events"
PROMO_KW = r"推廣費|公司推廣|Promotion Expense|推廣活動|Promotional"


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# inspect_galaxy_sjm_comp"]
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

    def basis(yr): return "pre" if yr == "25" else "post"

    # ── galaxy: comp-like 但唔喺 COMP_H ──
    L.append("\n" + "=" * 64 + "\n## galaxy：comp-like account 但唔喺 COMP_H（漏 → under）")
    g = df[df.e.eq("galaxy")]
    miss = g[g.ad.str.contains(COMP_KW, case=False, regex=True, na=False) & ~g.hid.isin(COMP_H)]
    for yr in ("23", "24", "25"):
        s = miss[miss.y2.eq(yr)]
        b = basis(yr)
        L.append(f"\n  [{yr}] comp-like 喺 COMP_H 外 Σ={s[b].sum():,.0f}萬（{b}）→ 現 H：")
        for (h, c, a), v in s.groupby(["hid", "code", "ad"])[b].sum().sort_values(key=lambda x: x.abs(), ascending=False).head(8).items():
            L.append(f"     {h:<14}{(c or '')[:9]:<9}{(a or '(空)')[:30]:<30}{v:>8,.0f}萬")

    # ── sjm: 推廣費 across 全 H ──
    L.append("\n" + "=" * 64 + "\n## sjm：推廣費(公司推廣活動)across 全 H（golden 當 comp；23 太少要撈）")
    s_ = df[df.e.eq("sjm") & df.ad.str.contains(PROMO_KW, case=False, regex=True, na=False)]
    for yr in ("23", "24", "25"):
        sub = s_[s_.y2.eq(yr)]
        b = basis(yr)
        L.append(f"\n  [{yr}] 推廣費 Σ={sub[b].sum():,.0f}萬（{b}）by 現 H：")
        for h, v in sub.groupby("hid")[b].sum().sort_values(key=lambda x: x.abs(), ascending=False).items():
            incomp = "✓comp" if h in COMP_H else "✗非comp"
            L.append(f"     {h:<16}{v:>9,.0f}萬  {incomp}")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
