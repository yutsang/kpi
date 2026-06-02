"""Wynn V review — per project: current vertical_label vs the project-team category column
(項目分類2 / 項目性質 = the '範疇第二層標籤' the project team fills) + amount, flag where they
imply different NG. Surfaces every remaining V mislabel so we can fix (override or column_map).

Run:
  python scripts/inspect_wynn_v.py --year 25
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parent.parent
# project-team theme (項目分類2 / 項目性質) → our V
THEME2V = {
    "博彩娛樂場場地的優化": "V_GAMING_VENUE", "博彩設施及設備的優化": "V_GAMING_EQUIP",
    "吸引外國客源": "V_INVITE_GUEST", "會議展覽": "V_MICE", "娛樂表演": "V_CONCERT",
    "體育盛事": "V_SPORT_EVENT", "文化藝術": "V_ART_EXHIBITION", "健康養生": "V_WELLNESS",
    "主題遊樂": "V_THEME_PARK", "美食之都": "V_RESTAURANT", "社區旅遊": "V_COMMUNITY",
    "海上旅遊": "V_MARITIME", "其他": "V_OTHER",
}
V2NG = {  # only need the NG of each V to compare buckets
    "V_GAMING_VENUE": "NG0", "V_GAMING_EQUIP": "NG0", "V_INVITE_GUEST": "NG1",
    "V_OVERSEAS_ROADSHOW": "NG1", "V_OVERSEAS_WEB_SEO": "NG1", "V_INVITE_AGENCY": "NG1",
    "V_MICE": "NG2", "V_CONCERT": "NG3", "V_SPORT_EVENT": "NG4", "V_VENUE_PERF_SPORT_MICE": "NG4",
    "V_ART_EXHIBITION": "NG5", "V_WELLNESS": "NG6", "V_THEME_PARK": "NG7",
    "V_RESTAURANT": "NG8", "V_FOOD_EVENT": "NG8", "V_COMMUNITY": "NG9",
    "V_MARITIME": "NG10", "V_OTHER": "NG11", "V_PROPERTY_UPGRADE": "NG11",
}


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--year", default="25")
    args = p.parse_args()
    pq = ROOT / "data/wynn/output/company_3_kpi_report.parquet"
    if not pq.exists():
        print(f"X {pq} missing"); return
    df = pd.read_parquet(pq)
    yc = next((c for c in ("report_period", "report_year", "years") if c in df.columns), None)
    if yc:
        df = df[df[yc].astype(str).str.startswith(args.year)].copy()

    proj = next((c for c in ("Name of Investment Project", "Sub project") if c in df.columns), None)
    cat = next((c for c in ("項目分類2", "項目性質", "範疇第二層標籤", "範疇") if c in df.columns), None)
    amtc = next((c for c in df.columns if "Amount" in str(c) or "Expense Amount" in str(c)), None)
    vlab = "vertical_label" if "vertical_label" in df.columns else None
    vid = "vertical_id" if "vertical_id" in df.columns else None
    print(f"cols → project={proj!r} category={cat!r} amount={amtc!r} vertical_id={vid!r}")
    if not proj or not vid:
        print(f"X missing project/vertical. cols={list(df.columns)}"); return
    df["_amt"] = pd.to_numeric(df[amtc], errors="coerce").fillna(0) if amtc else 0

    def _mode(s):
        s = s[s.astype(str).str.strip().ne("")]
        return s.value_counts().index[0] if len(s) else ""
    agg = df.groupby(df[proj].astype(str)).agg(
        cur_V=(vid, _mode), cur_Vlabel=(vlab or vid, _mode),
        team_cat=(cat, _mode) if cat else (vid, lambda s: ""),
        amount=("_amt", "sum"), rows=("_amt", "size")).reset_index().rename(columns={proj: "project"})
    agg["expect_V"] = agg["team_cat"].map(lambda t: THEME2V.get(str(t).strip(), ""))
    agg["cur_NG"] = agg["cur_V"].map(lambda v: V2NG.get(str(v), "?"))
    agg["expect_NG"] = agg["expect_V"].map(lambda v: V2NG.get(str(v), ""))
    agg["MISMATCH"] = (agg["expect_NG"] != "") & (agg["cur_NG"] != agg["expect_NG"])
    agg = agg.reindex(agg["amount"].abs().sort_values(ascending=False).index)

    out = ROOT / "results" / "wynn_v_review_25.xlsx"
    agg.to_excel(out, index=False)
    mm = agg[agg["MISMATCH"]]
    print(f"\nMISMATCH (cur NG ≠ project-team 類別 NG): {len(mm)} 項, Σ={mm['amount'].sum():,.0f}")
    print(f"{'project':12}{'cur_V':24}{'team_cat':16}{'→expect':22}{'amount':>15}")
    for _, r in mm.head(30).iterrows():
        print(f"{str(r['project'])[:11]:12}{str(r['cur_Vlabel'])[:22]:24}{str(r['team_cat'])[:14]:16}"
              f"{str(r['expect_V'])[:20]:22}{r['amount']:>15,.0f}")
    print(f"\n→ {out}  (full per-project review)")


if __name__ == "__main__":
    main()
