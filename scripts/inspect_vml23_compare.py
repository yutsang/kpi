"""inspect_vml23_compare.py — compare OUR VML-2023 H/V vs 項目組's own 分類1 (H) / 類別2 (V).

These live in the 2023 raw and survive into TAGGED_ROWS.parquet (NOT kpi_report.parquet, which
step5 curates down to mapped columns). REFERENCE ONLY — ours is the standard (memory taxonomy,
unified across 6 entities). Surfaces where we differ + the project↔subproject granularity.

  python scripts/inspect_vml23_compare.py --entity vml
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import pyarrow.parquet as pq
import yaml

ROOT = Path(__file__).resolve().parent.parent
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


def find(df, *subs):
    """first column whose stripped name contains ANY of subs (substring, robust to variants)."""
    for c in df.columns:
        cs = str(c).strip()
        if any(s in cs for s in subs):
            return c
    return None


def numify(s):
    return pd.to_numeric(s.astype(str).str.replace(",", "", regex=False).str.replace(r"^\s*-\s*$", "0", regex=True),
                         errors="coerce").fillna(0)


def best_amount(df, cands):
    """first candidate (in priority order) with a non-zero numeric sum — prefers the pipeline's
    cols['amount'] (= 調整後金額 for VML 23) over the raw gross 'Amount' / text columns."""
    for c in cands:
        if c and c in df.columns and numify(df[c]).abs().sum() > 0:
            return c
    return None


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--entity", default="vml")
    a = ap.parse_args()
    com = {"vml": "company_4"}.get(a.entity, "company_4")
    src = ROOT / "data" / a.entity / "interim" / f"{com}_tagged_rows.parquet"
    if not src.exists():
        print(f"X missing {src.relative_to(ROOT)} — run kedro first."); return
    cf = yaml.safe_load((ROOT / "conf" / com / "parameters.yml").open(encoding="utf-8"))
    cols = cf.get("columns", {}) or {}
    df = pq.read_table(src).replace_schema_metadata(None).to_pandas()
    if "report_period" in df.columns:
        df = df[df["report_period"].astype(str).str.startswith("23")].copy()

    # ---- column diagnostic so we SEE the real names ----
    print("=== columns containing 分類/類別/project/sub/名稱/ref (so we know the real names) ===")
    hit = [c for c in df.columns if any(k in str(c) for k in ("分類", "分类", "類別", "类别", "Sub", "sub", "Project", "project", "名稱", "ref", "Ref", "投資領域"))]
    print("   " + " | ".join(repr(str(c)) for c in hit))

    f1 = find(df, "分類1", "分类1", "分類 1")
    le2 = find(df, "類別2", "类别2", "類別 2", "類別2")
    amt = best_amount(df, [cols.get("amount"), "MOP Amt", "調整後金額", find(df, "Amount"), find(df, "金額")])
    df["_amt"] = numify(df[amt]) if amt else 0.0
    tot = df["_amt"].sum()
    print(f"\n[{a.entity} 23] rows={len(df):,}  amount_col={amt!r}")
    print(f"  resolved 項目組 cols → 分類1={f1!r}   類別2={le2!r}")

    print("\n===== TOTAL RECON vs 項目組 pivot =====")
    print(f"  our Σ          = {tot:,.0f}")
    print(f"  項目組 pivot    = {THEIR_TOTAL:,.0f}")
    print(f"  diff           = {tot - THEIR_TOTAL:,.0f}  ({(tot-THEIR_TOTAL)/THEIR_TOTAL*100 if THEIR_TOTAL else 0:+.2f}%)")

    if f1:
        df["_projH"] = df[f1].astype(str).str.strip().map(H_MAP).fillna("(unmapped)")
        df["_ourH"] = df.get("horizontal_id", pd.Series("", index=df.index)).astype(str)
        good = df[df["_projH"].ne("(unmapped)")]
        base = good["_amt"].abs().sum()
        ag = good.loc[good["_ourH"].eq(good["_projH"]), "_amt"].abs().sum()
        print(f"\n===== H COMPARE: our horizontal_id vs map(分類1)  [REFERENCE ONLY] =====")
        print(f"  $-weighted agreement: {ag:,.0f}/{base:,.0f} = {ag/base*100 if base else 0:.1f}%")
        d = good[good["_ourH"].ne(good["_projH"])].groupby(["_projH", "_ourH"])["_amt"].agg(["size", "sum"]).reset_index()
        d = d.reindex(d["sum"].abs().sort_values(ascending=False).index).head(20)
        print("  top diffs by Σamt (項目組_H → our_H):")
        for _, r in d.iterrows():
            print(f"     {r['_projH']:16s} -> {r['_ourH']:16s} Σ={r['sum']:>15,.0f} ({int(r['size'])} rows)")
        unm = df[df["_projH"].eq("(unmapped)")][f1].astype(str).str.strip().value_counts().head(12)
        unm = unm[unm.index != ""]
        if len(unm):
            print("  分類1 values NOT in H_MAP (add to map?): " + ", ".join(f"{v!r}×{n}" for v, n in unm.items()))
    else:
        print("\n  ⚠ 分類1 column not found — check the column-diagnostic line above for the real name.")

    if le2:
        vl = "vertical_label" if "vertical_label" in df.columns else "vertical_id"
        print(f"\n===== V COMPARE: our {vl} × 項目組 類別2  [granularity differs — cross-tab] =====")
        ct = df.groupby([le2, vl])["_amt"].sum().reset_index()
        ct = ct.reindex(ct["_amt"].abs().sort_values(ascending=False).index).head(25)
        for _, r in ct.iterrows():
            print(f"     類別2={str(r[le2])[:16]:18s} our_V={str(r[vl])[:18]:20s} Σ={r['_amt']:>15,.0f}")
    else:
        print("\n  ⚠ 類別2 column not found — check the diagnostic line.")

    pj = (cols.get("project") if cols.get("project") in df.columns else None) or find(df, "SubProject_Name")
    spc = find(df, "Subproject") or find(df, "項目名稱", "Sub project")
    print(f"\n===== project vs subproject granularity (drives 次數/unique count) =====")
    for nm, c in [("project", pj), ("subproject", spc)]:
        if c:
            s = df[c].astype(str).str.strip()
            bl = (s.eq("") | s.eq("nan")).mean() * 100
            print(f"  {nm:11s} col={c!r:22s} distinct={s.nunique():>6,}  blank={bl:4.0f}%{'  ⚠BLANK' if bl>30 else ''}")
        else:
            print(f"  {nm:11s} col=NOT FOUND ⚠")
    if pj and spc:
        x = df[pj].astype(str).str.strip(); y = df[spc].astype(str).str.strip()
        mism = df[x.ne(y) & y.ne("") & y.ne("nan")]
        print(f"  rows project != subproject: {len(mism):,}  (Σ={mism['_amt'].sum():,.0f})")
        for _, r in df.assign(x=x, y=y).groupby(["x", "y"]).size().reset_index(name="n").sort_values("n", ascending=False).head(10).iterrows():
            print(f"     proj={str(r['x'])[:20]:22s} | sub={str(r['y'])[:30]:32s} ×{int(r['n'])}")


if __name__ == "__main__":
    main()
