r"""inspect_wynn_buckets —— wynn 25/25_24SY/25_23SY bucket 對唔上 databook（user 手拉式）。

逐 bucket(25 / 25_24SY / 25_23SY) × metric(staff/comp/capex) dump 我哋金額（調整前+調整後），
畀 user 同佢 databook 手拉數逐 bucket 對。
Run（Windows）:  python scripts\inspect_wynn_buckets.py
出 results\inspect_wynn_buckets.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "inspect_wynn_buckets.txt"
COMP_H = {"H_HOTEL_ROOM", "H_VENUE", "H_FNB", "H_COMP_TICKET", "H_COMP_OTHER"}


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# inspect_wynn_buckets —— wynn 25 三 bucket 對 databook（萬）"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    e = df[df["entity"].astype(str).eq("wynn")].copy()
    e["pre"] = pd.to_numeric(e.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    e["post"] = pd.to_numeric(e.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    e["yb"] = e["year_bucket"].astype(str)
    e["hid"] = _s(e.get("horizontal_id", pd.Series("", index=e.index)))
    e["cap"] = _s(e.get("final_capex_opex", pd.Series("", index=e.index))).eq("Capex")
    e["staff"] = e["hid"].eq("H_LABOR") & ~e["cap"]
    e["comp"] = e["hid"].isin(COMP_H)

    L.append("\n  ── wynn by year_bucket：staff(opex) / comp / capex（調整前｜調整後）──")
    L.append(f"   {'bucket':<12}{'staff前':>9}{'staff後':>9}{'comp前':>9}{'comp後':>9}{'capex前':>10}{'capex後':>10}{'行':>8}")
    for yb in sorted(e["yb"].unique()):
        s = e[e.yb.eq(yb)]
        st = s[s.staff]; cm = s[s.comp]; cx = s[s.cap]
        L.append(f"   {yb:<12}{st['pre'].sum():>9,.0f}{st['post'].sum():>9,.0f}"
                 f"{cm['pre'].sum():>9,.0f}{cm['post'].sum():>9,.0f}"
                 f"{cx['pre'].sum():>10,.0f}{cx['post'].sum():>10,.0f}{len(s):>8,}")
    # 25 prefix 合計
    p25 = e[e.yb.str[:2].eq("25")]
    L.append(f"\n  25-prefix 合計：staff後={p25[p25.staff]['post'].sum():,.0f}（golden 10,539）"
             f" comp前={p25[p25.comp]['pre'].sum():,.0f}（golden 21,488）"
             f" capex前={p25[p25.cap]['pre'].sum():,.0f}")

    # 25_24SY / 25_23SY 係咩（睇 source/project，可能 wynn 根本唔應該有呢啲 SY）
    for yb in ("25_24SY", "25_23SY"):
        s = e[e.yb.eq(yb)]
        if s.empty:
            L.append(f"\n  {yb}: (冇行)"); continue
        L.append(f"\n  ── {yb}：{len(s):,}行 Σ前={s['pre'].sum():,.0f} —— 由邊嚟 ──")
        for col in ("source", "project", "account_desc"):
            if col in s.columns:
                g = s.groupby(_s(s[col]))["pre"].sum().sort_values(key=lambda x: x.abs(), ascending=False)
                g = g[g.index.astype(str) != ""]
                top = " | ".join(f"{str(k)[:24]}({v:,.0f})" for k, v in g.head(4).items())
                L.append(f"     {col:<12}: {top}")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
