r"""inspect_staff_dimensions —— 逐家搵「點拆 payroll 嘅 capex/opex」同「漏咗嘅資本化人工」。

問題兩個方向同一軸（payroll 嘅 capex/opex 維度）：
  A. MGM opex payroll 太高（capex人工=0）→ 邊個維度可以把部分推去 capex？
  B. galaxy/vml/sjm capex+opex payroll 太低 → 漏咗嘅資本化 payroll 匿喺邊？

每家 4 段：
  1. payroll 宇宙 by (final_capex_opex × horizontal_id) —— payroll 而家瞓喺邊
  2. payroll-keyword 行但唔係 H_LABOR —— 漏 tag / 掉咗去邊個 H（under-capture）
  3. per-NG capex:opex:payroll 矩陣 —— 邊個 NG 有大 capex 但 payroll 全 opex（可按項目 capex 比例拆）
  4. capex 行有 staff 訊號（項目組H=staff 或 payroll keyword）但唔係 H_LABOR —— 可補返嘅資本化人工

basis：調整前_萬（pre，= HQ Table 1 報告值口徑）；另показ 調整後 供 opex 對照。
Run（Windows）:  python scripts\inspect_staff_dimensions.py [ent ...]   (default 全6家)
出 results\inspect_staff_dimensions.txt  ← paste 返嚟
"""
from __future__ import annotations
import re
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "inspect_staff_dimensions.txt"
ENTS = sys.argv[1:] or ["galaxy", "wynn", "vml", "melco", "mgm", "sjm"]

KW = (r"salary|salaries|wage|payroll|overtime|provident|gratuity|severance|pension|"
      r"staff\s*cost|casual\s*(?:labou?r|worker)|contract\s*labou?r|headcount|"
      r"employee\s*(?:benefit|relation|transport|gift|training)|recruitment|bonus|"
      r"薪|工資|花紅|公積金|員工|人工|勞務")

# HQ Table 1（capex+opex 調整前，24-prefix = 24 + 24_23SY）
HQ_T1_24 = {"galaxy": 34926 + 16872, "wynn": 10219 + 9, "vml": 12163 + 12163,
            "melco": 14831 + 9, "mgm": 9270 + 3463, "sjm": 4115 + 2485}
HQ_OPEX_24 = {"galaxy": 14574, "wynn": 8765, "vml": 1457, "melco": 13912, "mgm": 6829, "sjm": 0}


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# inspect_staff_dimensions —— 逐家 payroll capex/opex 維度（萬，調整前 pre 為主，24-prefix）"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["pre"]  = pd.to_numeric(df.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    df["post"] = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    df["y2"] = df["year_bucket"].astype(str).str[:2]
    df["e"]  = df["entity"].astype(str)
    for c, n in [("horizontal_id", "hid"), ("final_capex_opex", "co"), ("account_code", "code"),
                 ("account_desc", "ad"), ("項目組H", "pth"), ("ng_code", "ng"), ("ng_label", "nglab"),
                 ("調整一級", "adj1"), ("調整二級", "adj2")]:
        df[n] = _s(df.get(c, pd.Series("", index=df.index)))
    df["code"] = df["code"].str.replace(r"\.0$", "", regex=True)
    df["cap"] = df["co"].eq("Capex")
    df["islab"] = df["hid"].eq("H_LABOR")
    df["kw"] = df["ad"].str.contains(KW, case=False, regex=True, na=False)
    df["pth_staff"] = df["pth"].str.contains(r"staff\s*cost|人工|薪|payroll", case=False, regex=True, na=False)

    for ent in ENTS:
        e = df[df.e.eq(ent) & df.y2.eq("24")]          # 24-prefix
        if not len(e):
            continue
        L.append("\n" + "█" * 84)
        L.append(f"█ {ent}  24-prefix   HQ-opex(Tab2)={HQ_OPEX_24.get(ent,0):,}  HQ-總(Tab1)={HQ_T1_24.get(ent,0):,}")
        L.append("█" * 84)

        # ── 1. payroll 宇宙：H_LABOR by capex/opex（payroll 瞓喺邊）──
        lab = e[e.islab]
        ox = lab[~lab.cap]["pre"].sum(); cx = lab[lab.cap]["pre"].sum()
        ox_post = lab[~lab.cap]["post"].sum()
        L.append(f"\n  ① H_LABOR：opex(pre)={ox:,.0f}  capex(pre)={cx:,.0f}  合計={ox+cx:,.0f}  [opex post={ox_post:,.0f} 對 HQ-opex]")
        L.append(f"     → 對 HQ：opex Δ(post)={ox_post-HQ_OPEX_24.get(ent,0):+,.0f} ｜ cap+opex Δ(pre)={ox+cx-HQ_T1_24.get(ent,0):+,.0f}")
        if cx == 0 and ox > 0:
            L.append(f"     ⚠ capex人工=0 → 全當 opex；若有資本化部分應拆走（搵 §3 大capex-NG）")

        # ── 2. payroll-keyword 但唔係 H_LABOR（under-capture：掉咗去邊）──
        miss = e[e.kw & ~e.islab]
        L.append(f"\n  ② payroll-keyword 但唔係 H_LABOR：{len(miss):,}行 Σ(pre)={miss['pre'].sum():,.0f}（漏 tag / 可能要拉返人工）")
        if len(miss):
            by_h = miss.groupby(["co", "hid"])["pre"].sum().sort_values(key=lambda s: s.abs(), ascending=False)
            for (co, h), v in by_h.head(6).items():
                L.append(f"       {co or '(空co)':<7}{h or '(空H)':<16}{v:>11,.0f}")
            L.append("     top account：")
            ga = miss.groupby(["cap", "code", "ad"])["pre"].sum().reset_index().sort_values("pre", key=lambda s: s.abs(), ascending=False)
            for _, r in ga.head(10).iterrows():
                tag = "CAP" if r["cap"] else "opx"
                L.append(f"       [{tag}] {(r['code'] or '(空)'):<10}{(r['ad'] or '(空)')[:34]:<34}{r['pre']:>10,.0f}")

        # ── 3. per-NG capex:opex:payroll 矩陣（拆 payroll 嘅維度）──
        L.append("\n  ③ per-NG  capex / opex / 其中payroll-opex / payroll-capex（搵大capex但payroll全opex嘅NG）")
        agg = e.groupby(["ng", "nglab"]).apply(
            lambda g: pd.Series({
                "capex": g[g.cap]["pre"].sum(),
                "opex": g[~g.cap]["pre"].sum(),
                "pay_ox": g[g.islab & ~g.cap]["pre"].sum(),
                "pay_cx": g[g.islab & g.cap]["pre"].sum(),
            })).reset_index()
        agg["capshare"] = agg["capex"] / (agg["capex"] + agg["opex"]).replace(0, 1)
        agg = agg.sort_values("capex", ascending=False)
        L.append(f"     {'NG':<6}{'capex':>10}{'opex':>10}{'pay_ox':>9}{'pay_cx':>9}{'cap%':>6}  nglabel")
        for _, r in agg.head(10).iterrows():
            L.append(f"     {r['ng']:<6}{r['capex']:>10,.0f}{r['opex']:>10,.0f}{r['pay_ox']:>9,.0f}{r['pay_cx']:>9,.0f}{r['capshare']*100:>5,.0f}%  {r['nglab'][:22]}")
        # 量化：若按 NG capshare 拆 payroll-opex → 會搬幾多去 capex
        movable = (agg["pay_ox"] * agg["capshare"]).sum()
        L.append(f"     → 若 payroll-opex 按各NG capex%拆：可搬 ~{movable:,.0f}萬 去 capex（opex 跌、cap+opex 不變）")

        # ── 4. capex 行有 staff 訊號但唔係 H_LABOR（漏咗嘅資本化人工）──
        cap_staff = e[e.cap & ~e.islab & (e.pth_staff | e.kw)]
        L.append(f"\n  ④ capex 行有 staff 訊號(項目組H=staff或keyword)但唔係H_LABOR：{len(cap_staff):,}行 Σ(pre)={cap_staff['pre'].sum():,.0f}")
        if len(cap_staff):
            gc = cap_staff.groupby(["code", "ad"])["pre"].sum().reset_index().sort_values("pre", key=lambda s: s.abs(), ascending=False)
            for _, r in gc.head(8).iterrows():
                L.append(f"       {(r['code'] or '(空)'):<10}{(r['ad'] or '(空)')[:34]:<34}{r['pre']:>10,.0f}")
            L.append(f"     → 呢啲若補入 capex-payroll：cap+opex Δ 收窄 ~{cap_staff['pre'].sum():,.0f}萬")

    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
