r"""investigate_staff —— staff cost (H_LABOR) 全面拆數 + tie 診斷。

3 個信號交叉：
  (A) 我哋 H_LABOR(opex)            ← recon 用緊嘅
  (B) 項目組H == staff               ← 項目組自己 H label = ground truth
  (C) account_desc keyword           ← 偵測「似人工但漏 tag」

輸出：
  0. per entity × year_bucket 我 H_LABOR + tie HQ（睇 SY 拆分）
  1. 項目組H top 值 + (項目組staff vs 我 vs HQ) 三方比
  2. UNDER（項目組=staff 但我冇 tag）/ OVER（我 tag 但項目組唔係）/ UNDER(keyword)
  3. H_LABOR 異常行：top 單行 + 空 desc（artifact）

Run（Windows）:  python scripts\investigate_staff.py
出 results\investigate_staff.txt  ← paste 返嚟
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
OUT = ROOT / "results" / "investigate_staff.txt"
ENTS = ["galaxy", "wynn", "vml", "melco", "mgm", "sjm"]
YBS = ["23", "24", "24_23SY", "25", "25_24SY", "25_23SY"]
HQ = {"galaxy": {"23": 17082, "24": 14574, "25": 23950}, "wynn": {"23": 6928, "24": 8765, "25": 10958},
      "vml": {"23": 790, "24": 1457, "25": 1253}, "melco": {"23": 7218, "24": 13912, "25": 27806},
      "mgm": {"23": 6357, "24": 6829, "25": 21467}, "sjm": {"23": 0, "24": 0, "25": 142}}
PTH_STAFF = r"staff|人工|薪|labor|labour|payroll|salary|wage"
KW = (r"salary|wage|payroll|overtime|provident|gratuity|severance|pension|staff\s*cost|"
      r"casual\s*worker|employee\s*(?:benefit|relation|transport)|recruitment\s*levy|bonus|"
      r"薪|工資|花紅|公積金|員工|人工")
L: list[str] = ["# investigate_staff —— staff cost 拆數 + tie 診斷（萬，recon basis）"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": "", "0.0": "0"})


def main():
    if not CSV.exists():
        L.append("!! tableau_combined_25.csv 揾唔到"); _w(); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    post = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    pre  = pd.to_numeric(df.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    yb = df["year_bucket"].astype(str)
    y2 = yb.str[:2]
    df["m"]  = post.where(y2.isin(["23", "24"]), pre)   # recon staff basis
    df["ry"] = y2                                         # report-year prefix: 23/24/25
    df["yb"] = yb
    df["e"]  = df["entity"].astype(str)
    hid   = _s(df.get("horizontal_id", pd.Series("", index=df.index)))
    co    = _s(df.get("final_capex_opex", pd.Series("", index=df.index)))
    desc  = _s(df.get("account_desc", pd.Series("", index=df.index)))
    descr = _s(df.get("description", pd.Series("", index=df.index)))
    acode = _s(df.get("account_code", pd.Series("", index=df.index))).str.replace(r"\.0$", "", regex=True)
    df["hid"] = hid; df["co"] = co; df["desc"] = desc; df["descr"] = descr; df["acode"] = acode
    df["pth"] = _s(df.get("項目組H", pd.Series("", index=df.index)))
    opex = ~co.eq("Capex")
    is_lab = hid.eq("H_LABOR")
    pth_staff = df["pth"].str.contains(PTH_STAFF, case=False, regex=True, na=False)
    kwmask = desc.str.contains(KW, case=False, regex=True, na=False)
    lab = df[is_lab & opex]

    # ── 0. per entity × bucket + tie ───────────────────────────────────────────
    L.append("=" * 82)
    L.append("## 0. 我 H_LABOR(opex) per entity × year_bucket（萬）+ tie HQ")
    for ent in ENTS:
        le = lab[lab.e.eq(ent)]
        b = {k: le[le.yb.eq(k)]["m"].sum() for k in YBS}
        o = {"23": b["23"], "24": b["24"] + b["24_23SY"], "25": b["25"] + b["25_24SY"] + b["25_23SY"]}
        h = HQ[ent]
        L.append(f"\n  {ent}")
        L.append("    buckets: " + " | ".join(f"{k}={b[k]:,.0f}" for k in YBS))
        L.append(f"    tie: 23 我{o['23']:,.0f} HQ{h['23']:,} Δ{o['23']-h['23']:+,.0f}"
                 f"  |  24(pref) 我{o['24']:,.0f} HQ{h['24']:,} Δ{o['24']-h['24']:+,.0f}"
                 f"  |  25(pref) 我{o['25']:,.0f} HQ{h['25']:,} Δ{o['25']-h['25']:+,.0f}")
    cap = df[is_lab & co.eq("Capex")].groupby("e")["m"].sum()
    L.append("\n  ── 註 capex H_LABOR（recon 唔計；under-tag 可能匿喺度）: "
             + " | ".join(f"{e}={cap.get(e, 0):,.0f}" for e in ENTS) + "（萬）")

    # ── 1. 項目組H reference + 三方比 ──────────────────────────────────────────
    L.append("\n" + "=" * 82)
    L.append("## 1. 項目組H（項目組自己 H label）top 值 + (項目組staff｜我｜HQ) 三方")
    for ent in ENTS:
        ee = df[df.e.eq(ent) & opex]
        nonblank = ee["pth"].ne("").mean() * 100
        L.append(f"\n  {ent}  (項目組H 非空 {nonblank:.0f}%)")
        top = ee.groupby("pth")["m"].sum().reset_index()
        top = top[top["pth"].ne("")].sort_values("m", key=lambda s: s.abs(), ascending=False).head(8)
        for _, r in top.iterrows():
            star = " ★staff" if re.search(PTH_STAFF, r["pth"], re.I) else ""
            L.append(f"     {r['pth'][:26]:<26}{r['m']:>13,.0f}萬{star}")
    L.append("\n  ── 三方（opex,萬）：項目組H=staff ｜ 我 H_LABOR ｜ HQ ──")
    L.append("   " + f"{'entity':<8}" + "".join(f"{'PT' + y:>11}{'我' + y:>11}{'HQ' + y:>11}" for y in ["23", "24", "25"]))
    for ent in ENTS:
        cells = []
        for y in ["23", "24", "25"]:
            ptv = df[df.e.eq(ent) & opex & pth_staff & df.ry.eq(y)]["m"].sum()
            ouv = lab[lab.e.eq(ent) & lab.ry.eq(y)]["m"].sum()
            cells.append(f"{ptv:>11,.0f}{ouv:>11,.0f}{HQ[ent][y]:>11,}")
        L.append(f"   {ent:<8}" + "".join(cells))
    L.append("   （PT≈HQ 但 我≠HQ → 係我 tag 同項目組唔同；PT≠HQ → 項目組 H label 唔等於 HQ staff basis）")

    # ── 2. UNDER / OVER ────────────────────────────────────────────────────────
    L.append("\n" + "=" * 82)
    L.append("## 2. 邊度 tag 唔同")
    under = df[opex & pth_staff & ~is_lab]
    L.append(f"\n  ── UNDER：項目組H=staff 但我唔係 H_LABOR(opex) = {len(under):,}行 / {under['m'].sum():,.0f}萬 ──")
    L.append("     （= 我可能漏 tag 嘅人工 → 解釋偏低）。我而家擺咗去邊個 hid：")
    for ent in ENTS:
        ue = under[under.e.eq(ent)]
        if abs(ue["m"].sum()) < 1:
            continue
        tops = ue.groupby("hid")["m"].sum().sort_values(key=lambda s: s.abs(), ascending=False).head(4)
        L.append(f"     {ent}: Σ={ue['m'].sum():,.0f}萬 → " + " | ".join(f"{h or '(空)'}={v:,.0f}" for h, v in tops.items()))
    g = under.groupby(["e", "acode", "desc"])["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False).head(20)
    L.append("\n     top account（under，似人工但漏 tag）：")
    for _, r in g.iterrows():
        L.append(f"     {r['e']:<8}{r['acode']:<10}{(r['desc'] or '(空)')[:32]:<32}{r['m']:>11,.0f}萬")

    over = lab[lab.pth.ne("") & ~lab.pth.str.contains(PTH_STAFF, case=False, regex=True, na=False)]
    L.append(f"\n  ── OVER：我 H_LABOR 但 項目組H≠staff（有 pth 值）= {len(over):,}行 / {over['m'].sum():,.0f}萬 ──")
    L.append("     （= 我可能 over-tag → 解釋偏高）。項目組覺得佢哋係咩：")
    g2 = over.groupby(["e", "pth", "acode", "desc"])["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False).head(20)
    for _, r in g2.iterrows():
        L.append(f"     {r['e']:<8}{r['pth'][:12]:<12}{r['acode']:<10}{(r['desc'] or '(空)')[:26]:<26}{r['m']:>11,.0f}萬")

    underkw = df[opex & kwmask & ~is_lab]
    L.append(f"\n  ── UNDER(keyword)：account_desc 似人工 但我唔係 H_LABOR = {len(underkw):,}行 / {underkw['m'].sum():,.0f}萬 ──")
    L.append("     （獨立於項目組H coverage；睇有冇大舊漏 tag）")
    L.append("   " + f"{'entity':<8}" + "".join(f"{y:>13}" for y in ["23", "24", "25"]))
    for ent in ENTS:
        cells = [f"{underkw[underkw.e.eq(ent) & underkw.ry.eq(y)]['m'].sum():>13,.0f}" for y in ["23", "24", "25"]]
        L.append(f"   {ent:<8}" + "".join(cells))

    # ── 3. anomaly ─────────────────────────────────────────────────────────────
    L.append("\n" + "=" * 82)
    L.append("## 3. H_LABOR(opex) 異常行")
    big = lab.reindex(lab["m"].abs().sort_values(ascending=False).index).head(15)
    L.append("\n  ── top 15 單行（|m|，揪出 MGM 2,421萬 嗰類大舊）──")
    for _, r in big.iterrows():
        d = (r["desc"] or r["descr"] or "(空)")[:34]
        L.append(f"     {r['e']:<8}{r['yb']:<9}{r['acode']:<10}{d:<34}{r['m']:>11,.0f}萬")
    bl = lab[lab.desc.eq("") & lab.descr.eq("")]
    L.append(f"\n  ── account_desc + description 都空白 H_LABOR Σ per entity（artifact 嫌疑）──")
    for ent in ENTS:
        be = bl[bl.e.eq(ent)]
        if abs(be["m"].sum()) >= 1:
            L.append(f"     {ent}: {len(be):,}行 / {be['m'].sum():,.0f}萬")
    _w()


def _w():
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
