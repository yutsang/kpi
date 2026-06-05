"""inspect_vml_compare.py — reconcile OUR VML H/V vs the project-team's OWN label columns, per year.
The team's columns differ by year (user-confirmed):
  23 (23JE): H = 分類1            V = 類別2
  24 (24JE): H/V among 分類1 (col DE) / 分類2 (col DB)  — NOTE 2024 has TWO 分類1 columns
  25 (25JE): H = 分類2            V = 進一步分類

So this tool FIRST prints every 分類/類別/進一步分類 column with its top values (so we SEE which axis it
is), THEN reconciles: our horizontal_id vs map(H col)  +  our vertical_label × (V col) cross-tab.
Override the picks with --hcol / --vcol once the values are confirmed.

  python scripts/inspect_vml_compare.py --year 25
  python scripts/inspect_vml_compare.py --year 24 --hcol "分類1" --vcol "分類2"
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
DEFAULTS = {"23": ("分類1", "類別2"), "24": ("分類1", "分類2"), "25": ("分類2", "進一步分類")}


def numify(s):
    return pd.to_numeric(s.astype(str).str.replace(",", "", regex=False).str.replace(r"^\s*-\s*$", "0", regex=True),
                         errors="coerce").fillna(0)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", required=True, choices=["23", "24", "25"])
    ap.add_argument("--hcol", default=None); ap.add_argument("--vcol", default=None)
    a = ap.parse_args()
    com = "company_4"
    src = ROOT / "data" / "vml" / "interim" / f"{com}_tagged_rows.parquet"
    if not src.exists():
        print(f"X missing {src.relative_to(ROOT)} — run kedro."); return
    cf = yaml.safe_load((ROOT / "conf" / com / "parameters.yml").open(encoding="utf-8"))
    cols = cf.get("columns", {}) or {}
    df = pq.read_table(src).replace_schema_metadata(None).to_pandas()
    if "report_period" in df.columns:
        df = df[df["report_period"].astype(str).str.startswith(a.year)].copy()
    amt = next((c for c in [cols.get("amount"), "MOP Amt", "調整後金額"] if c and c in df.columns
                and numify(df[c]).abs().sum() > 0), None)
    df["_amt"] = numify(df[amt]) if amt else 0.0
    print(f"[vml {a.year}] rows={len(df):,}  amount={amt!r}  Σ={df['_amt'].sum():,.0f}")

    # ---- show every candidate label column + its top values ----
    cand = [c for c in df.columns if any(k in str(c) for k in ("分類", "分类", "類別", "类别", "進一步", "进一步"))]
    print(f"\n=== candidate label columns ({len(cand)}) — top values (◀H if values are in H_MAP) ===")
    for c in cand:
        s = df[c].astype(str).str.strip().replace("nan", "")
        nz = s[s.ne("")]
        fill = len(nz) / len(df) * 100 if len(df) else 0
        vc = nz.value_counts().head(6)
        ishl = sum(1 for v in vc.index if v in H_MAP) >= max(1, len(vc) // 2)
        print(f"  {str(c)[:16]:18s} fill={fill:3.0f}%{' ◀H' if ishl else '   '}  " +
              ", ".join(f"{v[:10]}×{n}" for v, n in vc.items()))

    hcol = a.hcol or next((c for c in cand if c.strip() == DEFAULTS[a.year][0]), None)
    vcol = a.vcol or next((c for c in cand if c.strip() == DEFAULTS[a.year][1]), None)
    print(f"\nusing  H-col={hcol!r}  V-col={vcol!r}   (override with --hcol/--vcol)")

    # ---- H reconcile ----
    if hcol and hcol in df.columns:
        df["_projH"] = df[hcol].astype(str).str.strip().map(H_MAP).fillna("(unmapped)")
        df["_ourH"] = df.get("horizontal_id", pd.Series("", index=df.index)).astype(str)
        good = df[df["_projH"].ne("(unmapped)")]
        base = good["_amt"].abs().sum()
        ag = good.loc[good["_ourH"].eq(good["_projH"]), "_amt"].abs().sum()
        print(f"\n===== H: our horizontal_id vs map({hcol}) =====")
        print(f"  $-weighted agreement: {ag:,.0f}/{base:,.0f} = {ag/base*100 if base else 0:.1f}%")
        d = good[good["_ourH"].ne(good["_projH"])].groupby(["_projH", "_ourH"])["_amt"].agg(["size", "sum"]).reset_index()
        d = d.reindex(d["sum"].abs().sort_values(ascending=False).index).head(15)
        print("  top diffs (項目組_H → our_H):")
        for _, r in d.iterrows():
            print(f"     {r['_projH']:15s} -> {str(r['_ourH'])[:15]:15s} Σ={r['sum']:>14,.0f} ({int(r['size'])} rows)")
        unm = df[df["_projH"].eq("(unmapped)")][hcol].astype(str).str.strip().value_counts().head(10)
        unm = unm[unm.index != ""]
        if len(unm):
            print("  values NOT in H_MAP: " + ", ".join(f"{v!r}×{n}" for v, n in unm.items()))

    # ---- V reconcile (cross-tab) ----
    if vcol and vcol in df.columns:
        vl = "vertical_label" if "vertical_label" in df.columns else "vertical_id"
        print(f"\n===== V: our {vl} × 項目組 {vcol}  (top cells by Σamt) =====")
        ct = df.groupby([vcol, vl])["_amt"].sum().reset_index()
        ct = ct.reindex(ct["_amt"].abs().sort_values(ascending=False).index).head(25)
        for _, r in ct.iterrows():
            mark = "" if str(r[vcol]).strip() and str(r[vl]).strip() else ""
            print(f"     項目組={str(r[vcol])[:16]:18s} our_V={str(r[vl])[:18]:20s} Σ={r['_amt']:>14,.0f}")


if __name__ == "__main__":
    main()
