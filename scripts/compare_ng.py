"""ONE file comparing, per entity per NG: OUR pipeline total vs the project team's
ORIGINAL NG total — for 24 and 25, side by side.

OUR  = sum(amount) by ng_code derived from vertical_id (V_TO_NG).
THEIR= sum(amount) by the project team's raw NG column, ONLY when the entity has a
       column whose values are NG0–NG11 (e.g. Galaxy 'NG11 Category'). For entities
       whose project-team NG is not an NG0-11 column yet (Wynn/VML 項目性質,
       SJM/Melco 範疇 in project name), THEIR is left blank — tell me the source and
       I'll wire it.

Run (Windows):
  python scripts/compare_ng.py                       # all entities it can find
  python scripts/compare_ng.py --entity galaxy sjm wynn vml melco
Output: results/ng_compare.tsv   (one file, all entities stacked)
"""
from __future__ import annotations
import argparse, re, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import pyarrow.parquet as pq
import yaml

ROOT = Path(__file__).resolve().parent.parent
ENTITIES = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
            "vml": "company_4", "melco": "company_5", "mgm": "company_6"}
YEAR_CANDIDATES = ("report_period", "report_year", "Yr related", "years")
NG_LABEL = {"NG0": "博彩", "NG1": "吸引外國客源", "NG2": "會議展覽", "NG3": "娛樂表演",
            "NG4": "體育盛事", "NG5": "文化藝術", "NG6": "健康養生", "NG7": "主題遊樂",
            "NG8": "美食之都", "NG9": "社區旅遊", "NG10": "海上旅遊", "NG11": "其他"}
V_TO_NG = {
    "V_GAMING_VENUE": "NG0", "V_GAMING_EQUIP": "NG0",
    "V_OVERSEAS_OFFICE": "NG1", "V_OVERSEAS_WEB_SEO": "NG1", "V_OVERSEAS_ROADSHOW": "NG1",
    "V_INVITE_GUEST": "NG1", "V_INVITE_AGENCY": "NG1", "V_REGIONAL_TEAM": "NG1",
    "V_REGIONAL_SALES": "NG1", "V_PROMO_VIDEO": "NG1",
    "V_MICE": "NG2", "V_CONCERT": "NG3", "V_SPORT_EVENT": "NG4", "V_VENUE_PERF_SPORT_MICE": "NG4",
    "V_ART_EXHIBITION": "NG5", "V_MUSEUM": "NG5", "V_WELLNESS": "NG6", "V_THEME_PARK": "NG7",
    "V_RESTAURANT": "NG8", "V_FOOD_EVENT": "NG8", "V_COMMUNITY": "NG9", "V_MARITIME": "NG10",
    "V_FESTIVAL": "NG11", "V_PUBLIC_FACILITY": "NG11", "V_PROPERTY_UPGRADE": "NG11", "V_OTHER": "NG11",
}
NG_ORDER = [f"NG{i}" for i in range(12)]


def fuzzy(df, name):
    if not name: return None
    if name in df.columns: return name
    for c in df.columns:
        if str(c).strip() == str(name).strip(): return c
    return None


def year_mask(yr, tag):
    return yr.str.startswith(tag) | (yr == f"Yr 20{tag}") | (yr == f"20{tag}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", nargs="+", default=sorted(ENTITIES))
    args = ap.parse_args()

    out = ROOT / "results" / "ng_compare.tsv"
    out.parent.mkdir(parents=True, exist_ok=True)
    lines = ["entity\tng\tng_label\tour_25\tour_24\ttheir_25\ttheir_24\tdiff_25\tdiff_24"]
    for ent in args.entity:
        com = ENTITIES[ent]
        parquet = ROOT / "data" / ent / "output" / f"{com}_kpi_report.parquet"
        if not parquet.exists():
            print(f"[{ent}] skip — no {parquet.relative_to(ROOT)}"); continue
        cf = yaml.safe_load((ROOT / "conf" / com / "parameters.yml").open(encoding="utf-8"))
        cols = cf.get("columns", {}) or {}
        df = pq.read_table(parquet).replace_schema_metadata(None).to_pandas()
        amt = fuzzy(df, cols.get("amount", ""))
        df["_amt"] = pd.to_numeric(df[amt], errors="coerce").fillna(0) if amt else 0.0
        df["_our"] = df["vertical_id"].map(lambda v: V_TO_NG.get(str(v), "NG11"))
        # THEIR ng: only if a raw NG0-11 column exists
        their_col = None
        _ngc = fuzzy(df, cols.get("ng11_category", ""))
        if _ngc:
            tv = df[_ngc].astype("string").fillna("").str.strip().str.upper().str.replace(" ", "")
            if tv.str.fullmatch(r"NG\d+").fillna(False).mean() > 0.3:
                df["_their"] = tv.where(tv.str.fullmatch(r"NG\d+").fillna(False), "")
                their_col = _ngc
        ycol = next((c for c in YEAR_CANDIDATES if c in df.columns), None)
        yr = df[ycol].astype("string").fillna("")
        def tot(mask, col, ng):
            sub = df[mask & (df[col] == ng)]
            return float(sub["_amt"].sum())
        m25, m24 = year_mask(yr, "25"), year_mask(yr, "24")
        for ng in NG_ORDER:
            o25, o24 = tot(m25, "_our", ng), tot(m24, "_our", ng)
            if their_col:
                t25, t24 = tot(m25, "_their", ng), tot(m24, "_their", ng)
                d25, d24 = o25 - t25, o24 - t24
                tcell = (f"{t25:.0f}", f"{t24:.0f}", f"{d25:.0f}", f"{d24:.0f}")
            else:
                tcell = ("", "", "", "")
            lines.append(f"{ent}\t{ng}\t{NG_LABEL[ng]}\t{o25:.0f}\t{o24:.0f}\t"
                         + "\t".join(tcell))
        src = f"their=raw NG col {their_col!r}" if their_col else "their=(no NG0-11 col — tell me source)"
        print(f"[{ent}] our_25={df[m25]['_amt'].sum():,.0f}  our_24={df[m24]['_amt'].sum():,.0f}  {src}")
    out.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"\nwrote {out.relative_to(ROOT)}  ({len(args.entity)} entities)")


if __name__ == "__main__":
    main()
