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
    "capex": {"galaxy": {"23": 27449, "24": 101399, "25": 128629},
              "wynn": {"23": 75211, "24": 125999, "25": 141056},
              "vml": {"23": 70470, "24": 361174, "25": 143811},
              "melco": {"23": 35379, "24": 106400, "25": 127111},
              "mgm": {"23": 33644, "24": 129071, "25": 155597},
              "sjm": {"23": 26143, "24": 173735, "25": 157955}},
    "comp": {"galaxy": {"24": 55309, "25": 91878}, "wynn": {"24": 24459, "25": 21489},
             "vml": {"24": 6327, "25": 12802}, "melco": {"24": 6223, "25": 8424},
             "mgm": {"24": 10549, "25": 7914}, "sjm": {"24": 6553, "25": 3575}},
    "staff": {"galaxy": {"23": 17082, "24": 14574, "25": 23950}, "wynn": {"23": 6928, "24": 8765, "25": 10958},
              "vml": {"23": 790, "24": 1457, "25": 1253}, "melco": {"23": 7218, "24": 13912, "25": 27806},
              "mgm": {"23": 6357, "24": 6829, "25": 21467}, "sjm": {"23": 0, "24": 0, "25": 142}},
}
L: list[str] = ["# recon_hq —— 對 HQ Databook per-entity（萬；合併對數MOP=調整後23/24+調整前25）"]


def main():
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    post = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    pre = pd.to_numeric(df.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    yb = df["year_bucket"].astype(str)
    # EXACT bucket（唔好用 prefix；24_23SY ≠ 24）。合併對數MOP：23/24 bucket→調整後；25→調整前。SY bucket 唔計。
    df["m"] = post.where(yb.isin(["23", "24"]), pre)
    df["y2"] = yb   # exact bucket label
    co = df.get("final_capex_opex", pd.Series("", index=df.index)).astype(str).str.strip()
    hid = df.get("horizontal_id", pd.Series("", index=df.index)).astype(str).str.strip()
    df["ent"] = df["entity"].astype(str)

    def ours(ent, metric, yr):
        m = (df["ent"] == ent) & (df["y2"] == yr)
        if metric == "capex":
            m &= co.eq("Capex")
        elif metric == "comp":
            m &= hid.isin(COMP_H)
        elif metric == "staff":
            m &= hid.eq("H_LABOR") & (~co.eq("Capex"))   # staff = opex only（HQ 對數）
        return df.loc[m, "m"].sum()

    for metric in ("staff", "comp", "capex"):
        L.append(f"\n{'='*70}\n## {metric.upper()}（左 HQ｜右 我哋｜Δ=我哋−HQ）")
        yrs = ["23", "24", "25"] if metric != "comp" else ["24", "25"]
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
    _w()


def _w():
    out = ROOT / "results" / "recon_hq.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
