r"""inspect_melco23_sjm_staff —— 兩個 staff 疑點：
  A. melco 23 opex 人工 我=6,801 vs 新golden=4,172（+2,629 over）→ by account 睇邊度多咗
  B. sjm 人工全 capex、靠 is_labor 識別；user 話應 2億+(20,000萬+)，而家拉到好少
     → is_labor non-blank-non-0 嘅行：總額 / by capex-opex / 現H / by account（睇 2億喺邊、點解未 tag 人工）

Run（Windows）:  python scripts\inspect_melco23_sjm_staff.py
出 results\inspect_melco23_sjm_staff.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "inspect_melco23_sjm_staff.txt"


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# inspect_melco23_sjm_staff"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["pre"] = pd.to_numeric(df.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    df["post"] = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    df["e"] = df["entity"].astype(str)
    df["yb"] = df["year_bucket"].astype(str)
    df["y2"] = df["yb"].str[:2]
    for c, n in [("horizontal_id", "hid"), ("horizontal_label", "hl"), ("final_capex_opex", "co"),
                 ("account_code", "code"), ("account_desc", "ad"), ("is_labor", "il"), ("ng_code", "ng")]:
        df[n] = _s(df.get(c, pd.Series("", index=df.index)))
    df["code"] = df["code"].str.replace(r"\.0$", "", regex=True)
    df["cap"] = df["co"].eq("Capex")
    df["il_on"] = ~df["il"].isin(["", "0", "0.0", "N", "n", "False", "false", "FALSE", "No", "no"])

    # ── A. melco 23 opex 人工 ──
    m = df[df.e.eq("melco") & df.yb.eq("23") & df.hid.eq("H_LABOR") & ~df.cap]
    L.append(f"\n{'='*70}\n## A. melco 23 opex 人工(調整後)Σ={m['post'].sum():,.0f}萬  vs golden 4,172  Δ={m['post'].sum()-4172:+,.0f}")
    g = m.groupby(["code", "ad"])["post"].sum().reset_index().sort_values("post", key=lambda s: s.abs(), ascending=False)
    for _, r in g.head(18).iterrows():
        L.append(f"   {(r['code'] or ''):<12}{(r['ad'] or '(空)')[:40]:<40}{r['post']:>9,.0f}萬")
    L.append(f"\n   ── melco 23 全部 is_labor non-0 行（睇係咪有啲應剷）Σ(post)={df[df.e.eq('melco')&df.yb.eq('23')&df.il_on]['post'].sum():,.0f} ──")

    # ── B. sjm is_labor ──
    s = df[df.e.eq("sjm") & df.il_on]
    L.append(f"\n{'='*70}\n## B. sjm is_labor non-blank-non-0：{len(s):,}行  Σ(調整前)={s['pre'].sum():,.0f}萬  Σ(調整後)={s['post'].sum():,.0f}萬")
    L.append(f"   （user 話應 2億+ = 20,000萬+；而家 sjm 現 H_LABOR(任何)Σ(pre)={df[df.e.eq('sjm')&df.hid.eq('H_LABOR')]['pre'].sum():,.0f}萬）")
    L.append("\n   ── is_labor 行 by capex/opex × 現 H ──")
    for (cap, hid), v in s.groupby([s.cap, s.hid])["pre"].sum().sort_values(key=lambda x: x.abs(), ascending=False).head(10).items():
        L.append(f"     {'CAP' if cap else 'opx':<5}{(hid or '(空)'):<18}{v:>11,.0f}萬")
    L.append("\n   ── is_labor 行 by year ──")
    for yb, v in s.groupby("yb")["pre"].sum().items():
        L.append(f"     {yb:<10}{v:>11,.0f}萬")
    L.append("\n   ── is_labor 行 top account ──")
    g2 = s.groupby(["code", "ad"])["pre"].sum().reset_index().sort_values("pre", key=lambda x: x.abs(), ascending=False)
    for _, r in g2.head(15).iterrows():
        L.append(f"     {(r['code'] or ''):<12}{(r['ad'] or '(空)')[:38]:<38}{r['pre']:>11,.0f}萬")
    # 對照：sjm is_labor 但唔係 H_LABOR（= 漏 tag 嘅人工）
    miss = s[~s.hid.eq("H_LABOR")]
    L.append(f"\n   ── sjm is_labor 但唔係 H_LABOR（漏 tag）：{len(miss):,}行 Σ(pre)={miss['pre'].sum():,.0f}萬 → retag 可補返 ──")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
