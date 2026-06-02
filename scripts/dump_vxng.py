"""Dump each entity's 投資方向 (NG × Vertical × Σamount) for 25 and 24 to a TSV,
so it can be dropped into results/ for review. Mirrors build_master_audit_25's
source + NG logic exactly: reads data/<ent>/output/<code>_kpi_report.parquet,
derives ng_code from vertical_id (with the raw NG column override, e.g. Galaxy
'NG11 Category'), filters year on the same year column.

Run (Windows), after kedro run + (optionally) build_master_audit:
  python scripts/dump_vxng.py --entity galaxy
  python scripts/dump_vxng.py --entity galaxy sjm wynn melco   # several at once
Output: results/<ent>_vxng.tsv   (then copy those into results/ for Claude)
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
V_TO_NG = {  # vertical_id -> NG code (same mapping as build_master_audit_25)
    "V_GAMING_VENUE": "NG0", "V_GAMING_EQUIP": "NG0",
    "V_OVERSEAS_OFFICE": "NG1", "V_OVERSEAS_WEB_SEO": "NG1", "V_OVERSEAS_ROADSHOW": "NG1",
    "V_INVITE_GUEST": "NG1", "V_INVITE_AGENCY": "NG1", "V_REGIONAL_TEAM": "NG1",
    "V_REGIONAL_SALES": "NG1", "V_PROMO_VIDEO": "NG1",
    "V_MICE": "NG2", "V_CONCERT": "NG3", "V_SPORT_EVENT": "NG4", "V_VENUE_PERF_SPORT_MICE": "NG4",
    "V_ART_EXHIBITION": "NG5", "V_MUSEUM": "NG5", "V_WELLNESS": "NG6", "V_THEME_PARK": "NG7",
    "V_RESTAURANT": "NG8", "V_FOOD_EVENT": "NG8", "V_COMMUNITY": "NG9", "V_MARITIME": "NG10",
    "V_FESTIVAL": "NG11", "V_PUBLIC_FACILITY": "NG11", "V_PROPERTY_UPGRADE": "NG11", "V_OTHER": "NG11",
}


def fuzzy(df, name):
    if not name: return None
    if name in df.columns: return name
    for c in df.columns:
        if str(c).strip() == str(name).strip(): return c
    return None


def load_parquet(p):
    return pq.read_table(p).replace_schema_metadata(None).to_pandas()


def dump(ent: str):
    com = ENTITIES[ent]
    parquet = ROOT / "data" / ent / "output" / f"{com}_kpi_report.parquet"
    if not parquet.exists():
        print(f"[{ent}] X missing {parquet.relative_to(ROOT)} — run kedro first."); return
    cf = yaml.safe_load((ROOT / "conf" / com / "parameters.yml").open(encoding="utf-8"))
    cols = cf.get("columns", {}) or {}
    df = load_parquet(parquet)
    if "vertical_id" not in df.columns:
        print(f"[{ent}] X no vertical_id in parquet."); return

    amt = fuzzy(df, cols.get("amount", ""))
    df["_amt"] = pd.to_numeric(df[amt], errors="coerce").fillna(0) if amt else 0.0
    df["ng_code"] = df["vertical_id"].map(lambda v: V_TO_NG.get(str(v), "NG11"))
    # raw NG column override (project-team authoritative NG, e.g. Galaxy 'NG11 Category')
    ngc = fuzzy(df, cols.get("ng11_category", ""))
    if ngc:
        v = df[ngc].astype(str).str.strip().str.upper().str.replace(" ", "")
        isng = v.str.fullmatch(r"NG\d+").fillna(False)
        if isng.mean() > 0.3:
            df["ng_code"] = v.where(isng, df["ng_code"])
            print(f"[{ent}] NG from raw col {ngc!r} ({int(isng.sum()):,}/{len(df):,})")
    vlab = "vertical_label" if "vertical_label" in df.columns else "vertical_id"
    ycol = next((c for c in YEAR_CANDIDATES if c in df.columns), None)
    if not ycol:
        print(f"[{ent}] X no year column among {YEAR_CANDIDATES}."); return

    yr = df[ycol].astype("string").fillna("")

    def ngnum(s):
        m = re.search(r"NG(\d+)", str(s))
        return int(m.group(1)) if m else 99

    out = ROOT / "results" / f"{ent}_vxng.tsv"
    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("w", encoding="utf-8") as f:
        f.write(f"# {ent}  source={parquet.name}  year_col={ycol}\n")
        for tag in ("25", "24"):
            m = yr.str.startswith(tag) | (yr == f"Yr 20{tag}") | (yr == f"20{tag}")
            sub = df[m]
            tot = float(sub["_amt"].sum()) if not sub.empty else 0.0
            f.write(f"\n===== {tag}  TOTAL = {tot:,.0f}  ({len(sub):,} rows) =====\n")
            if sub.empty:
                continue
            g = sub.groupby(["ng_code", vlab])["_amt"].agg(["size", "sum"]).reset_index()
            g = g[g["sum"] != 0]
            # NG-only summary first (compact, for cross-entity sanity)
            ng = g.groupby("ng_code")["sum"].sum().reset_index()
            ng["_k"] = ng["ng_code"].map(ngnum)
            ng = ng.sort_values("_k")
            f.write("  NG summary:\n")
            for _, r in ng.iterrows():
                pct = (r["sum"] / tot * 100) if tot else 0
                f.write(f"    {r['ng_code']:<5} {r['sum']:>16,.0f}  {pct:4.1f}%\n")
            # NG x V detail (non-zero only)
            f.write("  NG x Vertical (non-zero):\n")
            f.write("  year\tng\tvertical_label\tn_rows\tamount\n")
            g["_k"] = g["ng_code"].map(ngnum)
            g = g.sort_values(["_k", "sum"], ascending=[True, False])
            for _, r in g.iterrows():
                f.write(f"  {tag}\t{r['ng_code']}\t{r[vlab]}\t{int(r['size'])}\t{r['sum']:,.0f}\n")
    print(f"[{ent}] wrote {out.relative_to(ROOT)}  ({len(df):,} rows)")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", nargs="+", required=True, choices=sorted(ENTITIES))
    for e in ap.parse_args().entity:
        dump(e)


if __name__ == "__main__":
    main()
