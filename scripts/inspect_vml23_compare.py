"""inspect_vml23_compare.py — compare OUR VML-2023 H/V vs 項目組's own 分類1 (H) / 類別2 (V),
which live in the 2023 raw and survive into kpi_report.parquet. REFERENCE ONLY — ours is the
standard (memory taxonomy, unified across 6 entities). This just surfaces where we differ and
checks the project↔subproject granularity that can distort counts/totals.

  python scripts/inspect_vml23_compare.py --entity vml

Reports: total recon vs their pivot (1,347,285,204); H agreement (our vs map(分類1)) + top $ diffs;
V cross-tab (our vertical_label × 類別2); project vs subproject distinct/blank/mismatch.
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import pyarrow.parquet as pq
import yaml

ROOT = Path(__file__).resolve().parent.parent
# 項目組 分類1 (H) -> our H taxonomy  (same map as VML conf column_map; for COMPARISON, not copy)
H_MAP = {
    "工程建設": "H_CONSTRUCTION", "拆除支出": "H_CONSTRUCTION", "設施採購": "H_EQUIP", "器具採購": "H_EQUIP",
    "娛樂表演合約成本": "H_PERFORMER", "專業服務費": "H_PROFESSIONAL", "代理費": "H_PROFESSIONAL",
    "人工成本": "H_LABOR", "管理成本分攤": "H_LABOR", "部門間成本分攤": "H_LABOR",
    "贊助費": "H_SPONSORSHIP", "公益性會展支出": "H_SPONSORSHIP", "授權費": "H_LICENSE",
    "媒體費用": "H_ADVERTISING", "廣告費": "H_ADVERTISING", "推廣費": "H_ADVERTISING",
    "租賃費": "H_LEASE", "食品飲料支出": "H_FNB", "會場支出": "H_VENUE", "客房支出": "H_HOTEL_ROOM",
    "租賃折扣": "H_DISCOUNT", "費用折扣": "H_DISCOUNT", "comp其他": "H_COMP_OTHER", "Comp其他": "H_COMP_OTHER",
    "交通費": "H_COMP_OTHER", "贈票支出": "H_COMP_TICKET", "維護維修": "H_MAINTENANCE", "其他": "H_OTHER",
}
THEIR_TOTAL = 1_347_285_204


def fuzzy(df, *names):
    for name in names:
        if not name:
            continue
        if name in df.columns:
            return name
        for c in df.columns:
            if str(c).strip() == str(name).strip():
                return c
    return None


def find_contains(df, *subs):
    for c in df.columns:
        if any(s in str(c) for s in subs):
            return c
    return None


def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", default="vml")
    a = ap.parse_args()
    ENT = {"vml": "company_4"}
    com = ENT.get(a.entity, "company_4")
    pqf = ROOT / "data" / a.entity / "output" / f"{com}_kpi_report.parquet"
    if not pqf.exists():
        print(f"X missing {pqf.relative_to(ROOT)}"); return
    cf = yaml.safe_load((ROOT / "conf" / com / "parameters.yml").open(encoding="utf-8"))
    cols = cf.get("columns", {}) or {}
    df = pq.read_table(pqf).replace_schema_metadata(None).to_pandas()
    yc = "report_period" if "report_period" in df.columns else None
    if yc:
        df = df[df[yc].astype(str).str.startswith("23")].copy()
    amt = fuzzy(df, cols.get("amount"), "調整後金額", "MOP Amt")
    df["_amt"] = pd.to_numeric(df[amt], errors="coerce").fillna(0)
    tot = df["_amt"].sum()

    print(f"\n[{a.entity} 23] rows={len(df):,}")
    print("===== TOTAL RECON vs 項目組 pivot =====")
    print(f"  our Σ調整後金額  = {tot:,.0f}")
    print(f"  項目組 pivot     = {THEIR_TOTAL:,.0f}")
    print(f"  diff            = {tot - THEIR_TOTAL:,.0f}   ({(tot-THEIR_TOTAL)/THEIR_TOTAL*100:+.2f}%)")

    f1 = find_contains(df, "分類1")
    le2 = find_contains(df, "類別2")
    print(f"\n  項目組 cols in parquet: 分類1={f1!r}  類別2={le2!r}")

    if f1:
        df["_projH"] = df[f1].astype(str).str.strip().map(H_MAP).fillna("(unmapped)")
        ourH = df.get("horizontal_id", pd.Series("", index=df.index)).astype(str)
        df["_ourH"] = ourH
        good = df[df["_projH"].ne("(unmapped)")]
        agree_amt = good.loc[good["_ourH"].eq(good["_projH"]), "_amt"].abs().sum()
        base_amt = good["_amt"].abs().sum()
        print(f"\n===== H COMPARE: our horizontal_id vs map(分類1)  (reference only) =====")
        print(f"  $-weighted agreement: {agree_amt:,.0f}/{base_amt:,.0f} = {agree_amt/base_amt*100 if base_amt else 0:.1f}%")
        diff = good[good["_ourH"].ne(good["_projH"])]
        d = diff.groupby(["_projH", "_ourH"])["_amt"].agg(["size", "sum"]).reset_index()
        d = d.reindex(d["sum"].abs().sort_values(ascending=False).index).head(20)
        print("  top differences by Σamt  (項目組_H → our_H):")
        for _, r in d.iterrows():
            print(f"     {r['_projH']:16s} -> {r['_ourH']:16s}  Σ={r['sum']:>15,.0f}  ({int(r['size'])} rows)")
        unm = df[df["_projH"].eq("(unmapped)")][f1].astype(str).value_counts().head(10)
        if len(unm):
            print(f"  分類1 values not in H_MAP: " + ", ".join(f"{v!r}×{n}" for v, n in unm.items()))

    if le2:
        vl = "vertical_label" if "vertical_label" in df.columns else "vertical_id"
        print(f"\n===== V COMPARE: our {vl} × 項目組 類別2  (granularity differs — cross-tab) =====")
        ct = df.groupby([le2, vl])["_amt"].sum().reset_index()
        ct = ct.reindex(ct["_amt"].abs().sort_values(ascending=False).index).head(25)
        for _, r in ct.iterrows():
            print(f"     類別2={str(r[le2])[:14]:16s} our_V={str(r[vl])[:16]:18s} Σ={r['_amt']:>15,.0f}")

    # project vs subproject granularity
    pj = fuzzy(df, cols.get("project"), "SubProject_Name")
    spc = fuzzy(df, "Subproject", "項目名稱")
    print(f"\n===== project vs subproject granularity =====")
    for nm, c in [("project", pj), ("subproject", spc)]:
        if c:
            s = df[c].astype(str).str.strip()
            blank = (s.eq("") | s.eq("nan")).mean() * 100
            print(f"  {nm:11s} col={c!r:20s} distinct={s.nunique():>6,}  blank={blank:4.0f}%{'  ⚠BLANK' if blank>30 else ''}")
        else:
            print(f"  {nm:11s} col=NOT FOUND ⚠")
    if pj and spc:
        a2 = df[pj].astype(str).str.strip(); b2 = df[spc].astype(str).str.strip()
        mism = df[a2.ne(b2) & b2.ne("") & b2.ne("nan")]
        print(f"  rows where project != subproject: {len(mism):,}  (Σ={mism['_amt'].sum():,.0f})")
        smp = mism.groupby([pj, spc]).size().reset_index(name="n").head(12)
        for _, r in smp.iterrows():
            print(f"     proj={str(r[pj])[:22]:24s} | sub={str(r[spc])[:30]:32s} ×{int(r['n'])}")


if __name__ == "__main__":
    main()
