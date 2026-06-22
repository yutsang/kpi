r"""inspect_wynn_detail —— wynn 兩件事：
  A. bucket 25/25_24SY/25_23SY total 對 databook（user：209,699/5,611/40，但 tableau 209,702/5,612/40）
     → dump 每 bucket total（MOP 未 round + 萬）+ by take_flag/netoff_flag/source，揾嗰 +3萬/+1萬 由邊嚟。
  B. wynn 24 opex staff 而家 8,765，databook 要 10,219（差 1,454）→ 揾 1,454 喺邊（似 staff 但落咗 專業/其他）。
Run（Windows）:  python scripts\inspect_wynn_detail.py
出 results\inspect_wynn_detail.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "inspect_wynn_detail.txt"
KW = r"staff|salary|salaries|wage|payroll|人工|薪|員工|support cost"


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# inspect_wynn_detail —— wynn bucket 精確 + 24 staff 1,454"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    e = df[df["entity"].astype(str).eq("wynn")].copy()
    e["amt"] = pd.to_numeric(e.get("amount_mop", 0), errors="coerce").fillna(0.0)
    e["post"] = pd.to_numeric(e.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    e["pre"] = pd.to_numeric(e.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    e["yb"] = e["year_bucket"].astype(str)
    for c, n in [("horizontal_id", "hid"), ("horizontal_label", "hl"), ("final_capex_opex", "co"),
                 ("account_desc", "ad"), ("take_flag", "tf"), ("netoff_flag", "nf"), ("source", "src")]:
        e[c if c in ("take_flag", "netoff_flag", "source") else n] = _s(e.get(c, pd.Series("", index=e.index)))
    e["tf"] = _s(e.get("take_flag", pd.Series("", index=e.index)))
    e["nf"] = _s(e.get("netoff_flag", pd.Series("", index=e.index)))
    e["src"] = _s(e.get("source", pd.Series("", index=e.index)))
    e["cap"] = e["co"].eq("Capex")

    # ── A. bucket total 精確 ──
    L.append("\n## A. wynn by year_bucket：total amount_mop（MOP + 萬）")
    L.append(f"   {'bucket':<12}{'MOP':>18}{'萬':>12}{'行':>9}")
    for yb in sorted(e["yb"].unique()):
        s = e[e.yb.eq(yb)]
        L.append(f"   {yb:<12}{s['amt'].sum():>18,.0f}{s['amt'].sum()/1e4:>12,.1f}{len(s):>9,}")
    for yb in ("25", "25_24SY", "25_23SY"):
        s = e[e.yb.eq(yb)]
        if s.empty:
            continue
        L.append(f"\n  ── {yb}（{s['amt'].sum()/1e4:,.1f}萬）by take_flag / netoff_flag（揾應剔嘅）──")
        L.append("     take_flag: " + " | ".join(f"{(k or '(空)')}={v/1e4:,.1f}萬" for k, v in s.groupby('tf')['amt'].sum().items()))
        L.append("     netoff_flag: " + " | ".join(f"{(k or '(空)')}={v/1e4:,.1f}萬" for k, v in s.groupby('nf')['amt'].sum().items()))
        # 細額尾數行（可能係嗰 3萬）
        small = s[s['amt'].abs() < 50000]   # < 5萬 嘅行
        L.append(f"     <5萬 細行：{len(small)}行 Σ={small['amt'].sum()/1e4:,.2f}萬")

    # ── B. wynn 24 staff 1,454 ──
    L.append("\n## B. wynn 24 opex：現 H_LABOR vs 似staff但落咗別處（揾 1,454）")
    o24 = e[e.yb.str.startswith("24") & ~e.cap]
    lab = o24[o24.hid.eq("H_LABOR")]
    L.append(f"   現 opex staff(H_LABOR) 24-prefix Σ後={lab['post'].sum():,.0f}（databook 要 10,219，差 {10219-lab['post'].sum():+,.0f}）")
    miss = o24[o24.ad.str.contains(KW, case=False, regex=True, na=False) & ~o24.hid.eq("H_LABOR")]
    L.append(f"\n   似 staff 但唔係 H_LABOR：{len(miss)}行 Σ後={miss['post'].sum():,.0f} → 而家落咗：")
    for (h, ad), v in miss.groupby(["hid", "ad"])["post"].sum().sort_values(key=lambda s: s.abs(), ascending=False).head(12).items():
        L.append(f"     {h:<15}{(ad or '(空)')[:34]:<34}{v:>9,.0f}萬")
    # wynn 全 opex by account top（睇 staff 應該係咩 account）
    L.append(f"\n   wynn 24 opex 全部 by account_desc top（睇 staff 結構）：")
    for ad, v in o24.groupby("ad")["post"].sum().sort_values(key=lambda s: s.abs(), ascending=False).head(12).items():
        inlab = o24[(o24.ad.eq(ad)) & o24.hid.eq("H_LABOR")]["post"].sum()
        L.append(f"     {(ad or '(空)')[:36]:<36}{v:>9,.0f}萬  (其中staff={inlab:,.0f})")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
