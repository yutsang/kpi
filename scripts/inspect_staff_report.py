r"""inspect_staff_report —— 6家 staff cost 對數 report（畀 user 回報用）。

三段：
  1. opex staff（H_LABOR & ~Capex）vs HQ Table 2（調整後）—— 對數主表 + 狀態
  2. capex/opex payroll 分佈 —— 揭示口徑（mgm capex payroll=0 = 全當 opex）
  3. 逐家 top account 組成（24 / 25 prefix）—— 解釋人工由咩科目砌成

basis（同 recon_hq 一致）：23=exact 調整後；24-prefix=調整後；25-prefix=調整前。
即 m = 調整後_萬 if year_bucket[:2] in (23,24) else 調整前_萬。

Run（Windows）:  python scripts\inspect_staff_report.py
出 results\inspect_staff_report.txt  ← paste / 回報
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "inspect_staff_report.txt"

ENTS = ["galaxy", "wynn", "vml", "melco", "mgm", "sjm"]

# HQ Table 2 —— opex 人工（調整後），對數金額 golden
HQ_OPEX = {
    "galaxy": {"23": 17082, "24": 14574, "25": 23950},
    "wynn":   {"23": 6928,  "24": 8765,  "25": 10958},
    "vml":    {"23": 790,   "24": 1457,  "25": 1253},
    "melco":  {"23": 7218,  "24": 13912, "25": 27806},
    "mgm":    {"23": 6357,  "24": 6829,  "25": 21467},
    "sjm":    {"23": 0,     "24": 0,     "25": 142},
}
# HQ Table 1 —— capex+opex 人工（調整前，report 值；23 / 24 / 24_23SY）
HQ_TOTAL = {
    "galaxy": {"23": 17082, "24": 34926, "24_23SY": 16872},
    "wynn":   {"23": 10219, "24": 10219, "24_23SY": 9},
    "vml":    {"23": 791,   "24": 12163, "24_23SY": 12163},
    "melco":  {"23": 15989, "24": 14831, "24_23SY": 9},
    "mgm":    {"23": 7231,  "24": 9270,  "24_23SY": 3463},
    "sjm":    {"23": 5554,  "24": 4115,  "24_23SY": 2485},
}
TOL = 800  # ±800萬 收貨門檻


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# Staff Cost 對數 report —— 6家 × {23,24,25}（萬）",
         "# basis: opex=H_LABOR&~Capex; 23=exact調整後, 24=prefix調整後, 25=prefix調整前；收貨=±800"]
    if not CSV.exists():
        L.append("!! tableau_combined_25.csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    post = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    pre  = pd.to_numeric(df.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    yb = df["year_bucket"].astype(str)
    y2 = yb.str[:2]
    df["m"]  = post.where(y2.isin(["23", "24"]), pre)
    df["yb"] = yb
    df["y2"] = y2
    df["e"]  = df["entity"].astype(str)
    df["hid"] = _s(df.get("horizontal_id", pd.Series("", index=df.index)))
    df["cap"] = _s(df.get("final_capex_opex", pd.Series("", index=df.index))).eq("Capex")
    df["code"] = _s(df.get("account_code", pd.Series("", index=df.index))).str.replace(r"\.0$", "", regex=True)
    df["adesc"] = _s(df.get("account_desc", pd.Series("", index=df.index)))
    lab = df[df["hid"].eq("H_LABOR")]

    def opex_staff(ent, yr):
        if yr == "23":
            m = lab[lab.e.eq(ent) & lab.yb.eq("23") & ~lab.cap]
        else:
            m = lab[lab.e.eq(ent) & lab.y2.eq(yr) & ~lab.cap]
        return m["m"].sum()

    def capex_staff(ent, yr):
        if yr == "23":
            m = lab[lab.e.eq(ent) & lab.yb.eq("23") & lab.cap]
        else:
            m = lab[lab.e.eq(ent) & lab.y2.eq(yr) & lab.cap]
        return m["m"].sum()

    # ── 1. opex staff vs HQ Table 2 ───────────────────────────────────────────
    L.append("\n" + "=" * 84)
    L.append("## 1. OPEX 人工 vs HQ Table 2（調整後）  ［Δ=我−HQ；✓=|Δ|≤800］")
    L.append("   " + f"{'entity':<8}" + "".join(f"{'HQ'+y:>9}{'我'+y:>9}{'Δ'+y:>8}{'':>2}" for y in ("23", "24", "25")))
    tot = {y: [0, 0] for y in ("23", "24", "25")}
    for ent in ENTS:
        cells = ""
        for y in ("23", "24", "25"):
            ours = opex_staff(ent, y)
            hq = HQ_OPEX.get(ent, {}).get(y, 0)
            d = ours - hq
            flag = "✓" if abs(d) <= TOL else ("🔴" if abs(d) > 1500 else "⚠")
            cells += f"{hq:>9,.0f}{ours:>9,.0f}{d:>8,.0f}{flag:>2}"
            tot[y][0] += hq; tot[y][1] += ours
        L.append("   " + f"{ent:<8}" + cells)
    tcells = "".join(f"{tot[y][0]:>9,.0f}{tot[y][1]:>9,.0f}{tot[y][1]-tot[y][0]:>8,.0f}{'':>2}" for y in ("23", "24", "25"))
    L.append("   " + f"{'總計':<8}" + tcells)

    # ── 2. capex / opex payroll 分佈 + STAFF_TOTAL vs HQ Table 1 ───────────────
    L.append("\n" + "=" * 84)
    L.append("## 2. capex+opex 人工 vs HQ Table 1（揭示資本化人工口徑；Δ=合計−HQ總）")
    L.append("   " + f"{'entity':<8}{'年':<6}{'opex人工':>10}{'capex人工':>11}{'合計':>10}{'HQTab1(總)':>12}{'Δ':>10}{'':>2}")
    for ent in ENTS:
        for y in ("23", "24"):
            ox = opex_staff(ent, y)
            cx = capex_staff(ent, y)
            # HQ Table1 prefix-24 = 24 + 24_23SY
            if y == "24":
                hqt = HQ_TOTAL.get(ent, {}).get("24", 0) + HQ_TOTAL.get(ent, {}).get("24_23SY", 0)
            else:
                hqt = HQ_TOTAL.get(ent, {}).get("23", 0)
            d = (ox + cx) - hqt
            flag = "✓" if abs(d) <= TOL else ("🔴" if abs(d) > 3000 else "⚠")
            L.append("   " + f"{ent:<8}{y:<6}{ox:>10,.0f}{cx:>11,.0f}{ox+cx:>10,.0f}{hqt:>12,.0f}{d:>10,.0f}{flag:>2}")
    L.append("   ⚠ capex人工=0 → 該家 payroll 全當 opex（冇拆資本化）；合計遠 under HQ總 = 資本化人工係 HQ top-down、唔喺 JE row")

    # ── 4. 逐家診斷摘要（回報用）──────────────────────────────────────────────
    NOTE = {
        "galaxy": "opex 24 −619/25 −153 收貨；23 −3,214 屬 23-build（blank-desc）。capex+opex 大幅 under（資本化人工 top-down 唔喺 row）。",
        "wynn":   "opex 24/25 全 tie。capex+opex under（資本化人工 top-down）。",
        "vml":    "opex 24 +435/25 −315 收貨。capex+opex under（80180 CONTRACT LABOR 入專業 + 資本化人工 top-down）。",
        "melco":  "opex 24 +263/25 +327 收貨。",
        "mgm":    "opex 24 +3,364：payroll 全當 opex(capex人工=0)，疑撈埋本應資本化嗰部分；25 +871。待 confirm payroll 表有冇 capex/opex 維度。",
        "sjm":    "opex≈0（HQ 25 得 142）；capex+opex under。",
    }

    # ── 3. 逐家 top account 組成 ──────────────────────────────────────────────
    L.append("\n" + "=" * 84)
    L.append("## 3. 逐家 opex 人工 top account（24 / 25 prefix，解釋組成）")
    for ent in ENTS:
        L.append(f"\n  ── {ent} ──")
        for y in ("24", "25"):
            sub = lab[lab.e.eq(ent) & lab.y2.eq(y) & ~lab.cap]
            if not len(sub):
                continue
            hq = HQ_OPEX.get(ent, {}).get(y, 0)
            L.append(f"    [{y}-prefix] 我Σ={sub['m'].sum():,.0f}  HQ={hq:,}  Δ={sub['m'].sum()-hq:+,.0f}")
            g = sub.groupby(["code", "adesc"])["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
            for _, r in g.head(8).iterrows():
                L.append(f"       {(r['code'] or '(空)'):<11}{(r['adesc'] or '(空)')[:38]:<38}{r['m']:>10,.0f}")

    # ── 4. 逐家診斷摘要 ───────────────────────────────────────────────────────
    L.append("\n" + "=" * 84)
    L.append("## 4. 逐家診斷摘要（回報用）")
    for ent in ENTS:
        L.append(f"  • {ent}: {NOTE.get(ent, '')}")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste / 回報")


if __name__ == "__main__":
    main()
