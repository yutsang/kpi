r"""對 HQ Databook（per-entity）—— capex / comp / staff(opex only)。
合併對數MOP = 調整後(23/24 bucket) + 調整前(25 bucket)；HQ 25 = 報告 = 調整前。

HQ 數（user 2026-06-19 貼，順序 galaxy/wynn/vml/melco/mgm/sjm）：
  capex 23/24/25、comp 24/25、staff(opex) 23/24/25。
我哋：capex = final_capex_opex==Capex；comp = comp H 五類；staff = H_LABOR(已改 opex-only)。

Run（Windows）:  python scripts\recon_hq.py
出 results\recon_hq.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
ENTS = ["galaxy", "wynn", "vml", "melco", "mgm", "sjm"]
COMP_H = {"H_HOTEL_ROOM", "H_VENUE", "H_FNB", "H_COMP_TICKET", "H_COMP_OTHER"}

HQ = {
    "capex": {"galaxy": {"23": 27449, "24": 117129, "25": 131838},  # 25=131838 (2026-06-24 換項目組adj+Capex/Opex欄重分類後 golden)
              "wynn": {"23": 75211, "24": 125999, "25": 141056},
              "vml": {"23": 70470, "24": 361175, "25": 135760},
              "melco": {"23": 35281, "24": 106029, "25": 127111},  # 23=35281 user 2026-06-23 raw+67M入capex(intended)
              "mgm": {"23": 33644, "24": 129071, "25": 155597},
              "sjm": {"23": 26143, "24": 173735, "25": 157955}},
    # comp golden（user 2026-06-22 新表：23/24=調整後, 25=報告(調整前)）
    "comp": {"galaxy": {"23": 32204, "24": 55321, "25": 91883}, "wynn": {"23": 14232, "24": 24459, "25": 21488},
             "vml": {"23": 6173, "24": 6335, "25": 9145}, "melco": {"23": 3247, "24": 6374, "25": 8424},  # vml25=9145(項目組已改折扣，照舊)
             "mgm": {"23": 10127, "24": 9073, "25": 7914}, "sjm": {"23": 2970, "24": 6092, "25": 3575}},
    # staff golden（user 2026-06-22 新表：23/24=調整後, 25=報告）
    "staff": {"galaxy": {"23": 17082, "24": 14574, "25": 23950}, "wynn": {"23": 6928, "24": 8765, "25": 10539},  # wynn24=8765(opex;10219嗰個capex+opex不準確,棄)
              "vml": {"23": 790, "24": 1457, "25": 790}, "melco": {"23": 4172, "24": 14563, "25": 27806},
              "mgm": {"23": 6357, "24": 6829, "25": 21467}, "sjm": {"23": 0, "24": 0, "25": 142}},
}
# Table 1 = capex+opex 人工成本 調整前（user 2026-06-19），exact bucket 23/24/24_23SY
STAFF_TOTAL = {"galaxy": {"23": 17082, "24": 34926, "24_23SY": 16872}, "sjm": {"23": 5554, "24": 4115, "24_23SY": 2485},
               "wynn": {"23": 10219, "24": 10219, "24_23SY": 9}, "vml": {"23": 791, "24": 12163, "24_23SY": 12163},
               "melco": {"23": 15989, "24": 14831, "24_23SY": 9}, "mgm": {"23": 7231, "24": 9270, "24_23SY": 3463}}
L: list[str] = ["# recon_hq —— 對 HQ Databook per-entity（萬；合併對數MOP=調整後23/24+調整前25）"]


def main():
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    post = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    pre = pd.to_numeric(df.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    yb = df["year_bucket"].astype(str)
    y2 = yb.str[:2]
    # capex/staff: 23/24-prefix=調整後, 25=調整前 (m_pref)
    # comp basis（user 2026-06-22 新表）: 23+24=調整後, 25=報告(調整前)
    #   → m_comp = pre.where(y2=="25", post)  只 y2=="25" 用 調整前
    df["bk_exact"] = yb
    df["bk_pref"] = y2
    df["m_exact"] = post.where(yb.isin(["23", "24"]), pre)
    df["m_pref"] = post.where(y2.isin(["23", "24"]), pre)
    df["m_comp"] = pre.where(y2.eq("25"), post)   # comp: 25 → 調整前(報告); 23/24 → 調整後
    df["pre_amt"] = pre
    co = df.get("final_capex_opex", pd.Series("", index=df.index)).astype(str).str.strip()
    hid = df.get("horizontal_id", pd.Series("", index=df.index)).astype(str).str.strip()
    df["ent"] = df["entity"].astype(str)

    def ours(ent, metric, yr):
        # row filter: 23=exact "23"; 24/25 = spent-year prefix
        if yr == "23":
            m = (df["ent"] == ent) & (df["bk_exact"] == "23")
        else:
            m = (df["ent"] == ent) & (df["bk_pref"] == yr)
        if metric == "capex":
            m &= co.eq("Capex")
            return df.loc[m, "m_pref"].sum()
        elif metric == "comp":
            m &= hid.isin(COMP_H)
            return df.loc[m, "m_comp"].sum()   # comp-specific basis
        elif metric == "staff":
            m &= hid.eq("H_LABOR") & (~co.eq("Capex"))
            return df.loc[m, "m_pref"].sum()

    for metric in ("staff", "comp", "capex"):
        L.append(f"\n{'='*70}\n## {metric.upper()}（左 HQ｜右 我哋｜Δ=我哋−HQ）")
        yrs = ["23", "24", "25"]
        hdr = "   " + f"{'entity':<8}" + "".join(f"{'HQ'+y:>10}{'我'+y:>10}{'Δ'+y:>9}" for y in yrs)
        L.append(hdr)
        tot = {y: [0, 0] for y in yrs}
        for ent in ENTS:
            cells = []
            for y in yrs:
                h = HQ[metric].get(ent, {}).get(y, 0)
                o = ours(ent, metric, y)
                tot[y][0] += h; tot[y][1] += o
                cells.append(f"{h:>10,.0f}{o:>10,.0f}{o-h:>9,.0f}")
            L.append(f"   {ent:<8}" + "".join(cells))
        L.append(f"   {'總計':<8}" + "".join(f"{tot[y][0]:>10,.0f}{tot[y][1]:>10,.0f}{tot[y][1]-tot[y][0]:>9,.0f}" for y in yrs))

    # STAFF_TOTAL：capex+opex 人工成本 調整前 vs Table 1（exact bucket 23/24/24_23SY）
    L.append(f"\n{'='*70}\n## STAFF_TOTAL（capex+opex 人工 調整前 vs Table 1）")
    yrs2 = ["23", "24", "24_23SY"]
    L.append("   " + f"{'entity':<8}" + "".join(f"{'HQ'+y:>10}{'我'+y:>10}{'Δ'+y:>9}" for y in yrs2))
    for ent in ENTS:
        cells = []
        for y in yrs2:
            h = STAFF_TOTAL.get(ent, {}).get(y, 0)
            m2 = (df["ent"] == ent) & (df["bk_exact"] == y) & hid.eq("H_LABOR")
            o = df.loc[m2, "m_exact"].sum()   # 調整後（同 opex recon basis；Table 1 = 報告值）
            cells.append(f"{h:>10,.0f}{o:>10,.0f}{o-h:>9,.0f}")
        L.append(f"   {ent:<8}" + "".join(cells))
    _w()


def _w():
    out = ROOT / "results" / "recon_hq.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
