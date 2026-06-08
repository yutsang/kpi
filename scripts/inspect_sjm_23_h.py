"""SJM 2023 H over-concentration diagnosis — 廣告及推廣 58% + 建設與設施 40% = 98% of the year.

Breaks down the two mega-buckets to see WHY they're so big:
  - by account_code (Cost Element) → which accounts drive them
  - by horizontal_source (rule:* vs llm vs row-override) → rule-collapse vs LLM-default
  - by capex/opex → how much 建設 is capex_force'd (any capex H∉{construction,equip,labor}→建設)
  - account_desc top values → is 廣告 firing on a few 推廣/媒體 accts (legit) or scattered (collapse)
Also lists, per NG, the dominant H — to spot NGs that are 100% one bucket (e.g. NG4 體育 all 廣告).

Run (Windows):  python scripts/inspect_sjm_23_h.py
Output: prints + results/inspect_sjm_23_h.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
TR = ROOT / "data" / "sjm" / "interim" / "company_2_tagged_rows.parquet"
AMT = "Val/COArea Crcy"
AC = "Cost Element"
AD = "Cost element descr."
TARGETS = ["建設與設施支出", "廣告及推廣"]


def main():
    L = ["# inspect_sjm_23_h — why 建設 + 廣告 dominate"]
    if not TR.exists():
        L.append(f"X {TR} missing"); _w(L); return
    df = pd.read_parquet(TR)
    per = next((c for c in ("report_period", "report_year", "years") if c in df.columns), None)
    if per:
        df = df[df[per].astype(str).str.startswith("23")].copy()
    a = pd.to_numeric(df[AMT], errors="coerce").fillna(0.0)
    tot = a.abs().sum() or 1
    hl = df["horizontal_label"].astype("string").fillna("(blank)")
    L.append(f"\n23 rows={len(df):,}  Σ|amt|={a.abs().sum():,.0f}")

    # overall H distribution
    L.append("\n## H distribution (Σ|amt| M / %):")
    for v, x in a.abs().groupby(hl).sum().sort_values(ascending=False).items():
        L.append(f"   {str(v)[:18]:18s} {x/1e6:9.2f}M  {x/tot*100:5.1f}%")

    # capex/opex × H
    co = df["final_capex_opex"].astype("string").fillna("") if "final_capex_opex" in df.columns else None
    if co is not None:
        L.append("\n## final_capex_opex × H (Σ|amt| M):")
        ct = pd.crosstab(co, hl, values=a.abs(), aggfunc="sum").fillna(0) / 1e6
        L.append(ct.round(1).to_string())

    src = df["horizontal_source"].astype("string").fillna("") if "horizontal_source" in df.columns else None
    acc = df[AC].astype("string").fillna("").str.strip() if AC in df.columns else None
    add = df[AD].astype("string").fillna("").str.strip() if AD in df.columns else None

    for tgt in TARGETS:
        m = hl.eq(tgt)
        if not m.any():
            continue
        L.append(f"\n{'='*70}\n## {tgt}: {int(m.sum()):,} rows  Σ|amt|={a.abs()[m].sum()/1e6:.1f}M")
        if src is not None:
            L.append("   horizontal_source: " + str(a.abs()[m].groupby(src[m]).sum().round(0).sort_values(ascending=False).head(6).to_dict()))
        if co is not None:
            L.append("   capex/opex: " + str(a.abs()[m].groupby(co[m]).sum().round(0).to_dict()))
        if acc is not None:
            L.append("   -- top account_code (Cost Element) by Σ|amt| --")
            g = a.abs()[m].groupby(acc[m]).agg(["sum", "count"]).sort_values("sum", ascending=False)
            for code, row in g.head(15).iterrows():
                L.append(f"      {str(code)[:30]:30s} {row['sum']/1e6:8.2f}M  ({int(row['count'])} rows)")
        if add is not None:
            L.append("   -- top account_desc by Σ|amt| --")
            g = a.abs()[m].groupby(add[m]).sum().sort_values(ascending=False)
            for d, x in g.head(12).items():
                L.append(f"      {str(d)[:40]:40s} {x/1e6:8.2f}M")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_sjm_23_h.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
