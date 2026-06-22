r"""inspect_vml_discount —— 揾 VML 25 折扣優惠喺我哋 data 邊度（HQ vml25 comp=12,802 含折扣，我哋 comp≈9,169）。

gap ≈ 3,633 折扣。睇佢而家 tag 咗做咩 H、account_desc 係咩 → 決定點加返落 comp。
Run（Windows）:  python scripts\inspect_vml_discount.py
出 results\inspect_vml_discount.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "inspect_vml_discount.txt"
COMP_H = {"H_HOTEL_ROOM", "H_VENUE", "H_FNB", "H_COMP_TICKET", "H_COMP_OTHER"}
DISC = r"折扣|優惠|discount|rebate|allowance|減免|讓利|coupon|comp|贈|complimentary"


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# inspect_vml_discount —— VML 25 折扣喺邊（HQ comp=12,802含折扣 vs 我≈9,169）"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["pre"] = pd.to_numeric(df.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    e = df[(df["entity"].astype(str).eq("vml")) & (df["year_bucket"].astype(str).str[:2].eq("25"))].copy()
    for c, n in [("horizontal_id", "hid"), ("horizontal_label", "hl"), ("account_code", "code"),
                 ("account_desc", "ad"), ("comp_type", "ct")]:
        e[n] = _s(e.get(c, pd.Series("", index=e.index)))
    e["code"] = e["code"].str.replace(r"\.0$", "", regex=True)
    comp_now = e[e.hid.isin(COMP_H)]["pre"].sum()
    L.append(f"\n  VML 25 現 comp(COMP_H 五類) Σ(調整前)={comp_now:,.0f}萬  | HQ=12,802 → gap≈{comp_now-12802:,.0f}")

    # ── 非 COMP_H 但 account_desc/comp_type 似折扣/comp ──
    nonc = e[~e.hid.isin(COMP_H)]
    disc = nonc[nonc["ad"].str.contains(DISC, case=False, regex=True, na=False)
                | nonc["ct"].str.contains(DISC, case=False, regex=True, na=False)]
    L.append(f"\n  ── 非comp-H 但似折扣/comp（{len(disc):,}行 Σ={disc['pre'].sum():,.0f}萬）by 現H × account_desc ──")
    g = disc.groupby(["hid", "code", "ad"])["pre"].sum().reset_index().sort_values("pre", key=lambda s: s.abs(), ascending=False)
    for _, r in g.head(20).iterrows():
        L.append(f"     {r['hid']:<15}{(r['code'] or ''):<10}{(r['ad'] or '(空)')[:34]:<34}{r['pre']:>9,.0f}")

    # ── comp_type 分佈（睇有冇「折扣」類 comp_type 但落咗非comp-H）──
    L.append(f"\n  ── VML25 comp_type 分佈（全部，睇有冇折扣類）──")
    for ct, amt in e.groupby("ct")["pre"].sum().sort_values(key=lambda s: s.abs(), ascending=False).head(15).items():
        in_comp = e[(e.ct.eq(ct)) & e.hid.isin(COMP_H)]["pre"].sum()
        L.append(f"     ct={(ct or '(空)')[:24]:<24} Σ={amt:>9,.0f}  其中已在comp-H={in_comp:,.0f}")

    # ── 大額非comp-H account（top，人眼睇邊個係折扣）──
    L.append(f"\n  ── VML25 非comp-H top account（睇 3,633 折扣匿喺邊）──")
    g2 = nonc.groupby(["hid", "code", "ad"])["pre"].sum().reset_index().sort_values("pre", key=lambda s: s.abs(), ascending=False)
    for _, r in g2.head(18).iterrows():
        L.append(f"     {r['hid']:<15}{(r['code'] or ''):<10}{(r['ad'] or '(空)')[:34]:<34}{r['pre']:>9,.0f}")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
