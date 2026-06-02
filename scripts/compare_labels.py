"""對數工具:OUR 分類（vertical_label / horizontal_label）vs 項目組嘅 label 欄,
睇 row-level 差異 + crosstab。Galaxy 用 基礎|一級標簽 / 基礎|二級標簽;其他 entity 可指定欄。

Run (Windows):
  python scripts/compare_labels.py --entity galaxy --their "基礎|一級標簽" --our horizontal_label --year 25
  python scripts/compare_labels.py --entity galaxy --their "基礎|二級標簽" --our horizontal_label --year 25 --rows
Output:
  results/<ent>_<year>_<their>_vs_<our>.tsv         — crosstab (their x our, Σamount)
  results/<ent>_<year>_<their>_vs_<our>_rows.tsv    — row-level (only with --rows)
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


def fuzzy(df, name):
    if not name: return None
    if name in df.columns: return name
    low = str(name).strip().lower()
    for c in df.columns:
        if str(c).strip().lower() == low: return c
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", required=True, choices=sorted(ENTITIES))
    ap.add_argument("--their", required=True, help="項目組 label 欄名 (e.g. 基礎|一級標簽)")
    ap.add_argument("--our", default="horizontal_label", help="horizontal_label | vertical_label")
    ap.add_argument("--year", default=None)
    ap.add_argument("--rows", action="store_true", help="埋多一個 row-level TSV")
    args = ap.parse_args()
    ent, com = args.entity, ENTITIES[args.entity]

    parquet = ROOT / "data" / ent / "output" / f"{com}_kpi_report.parquet"
    if not parquet.exists():
        print(f"X missing {parquet.relative_to(ROOT)}"); return
    cf = yaml.safe_load((ROOT / "conf" / com / "parameters.yml").open(encoding="utf-8"))
    cols = cf.get("columns", {}) or {}
    df = pq.read_table(parquet).replace_schema_metadata(None).to_pandas()

    their = fuzzy(df, args.their)
    if not their:
        print(f"X column {args.their!r} not in parquet. cols sample:")
        for c in df.columns: print("   ", c)
        return
    if args.our not in df.columns:
        print(f"X our column {args.our!r} not in parquet."); return

    ycol = next((c for c in YEAR_CANDIDATES if c in df.columns), None)
    if args.year and ycol:
        yr = df[ycol].astype("string").fillna("")
        df = df[yr.str.startswith(str(args.year)) | (yr == f"Yr 20{args.year}") | (yr == f"20{args.year}")].copy()
    amt = fuzzy(df, cols.get("amount", ""))
    df["_amt"] = pd.to_numeric(df[amt], errors="coerce").fillna(0) if amt else 0.0
    df["_their"] = df[their].astype("string").fillna("").str.strip().replace({"nan": ""})
    df["_our"] = df[args.our].astype("string").fillna("").str.strip()

    tag = re.sub(r"[^\w]+", "_", f"{args.their}_vs_{args.our}")[:50]
    yr_s = args.year or "all"
    out = ROOT / "results" / f"{ent}_{yr_s}_{tag}.tsv"
    out.parent.mkdir(parents=True, exist_ok=True)

    # per their-label: total + breakdown by our (so divergence is obvious)
    with out.open("w", encoding="utf-8") as f:
        f.write(f"# {ent} {yr_s}  their={their!r}  vs  our={args.our!r}\n")
        f.write("their_label\ttheir_total\tour_label\tn_rows\tamount\tshare%\n")
        for tl, sub in sorted(df.groupby("_their"), key=lambda kv: -kv[1]["_amt"].sum()):
            tot = sub["_amt"].sum()
            if tot == 0 and len(sub) == 0: continue
            g = sub.groupby("_our")["_amt"].agg(["size", "sum"]).reset_index().sort_values("sum", ascending=False)
            for _, r in g.iterrows():
                if r["sum"] == 0: continue
                sh = (r["sum"] / tot * 100) if tot else 0
                f.write(f"{tl or '(空)'}\t{tot:,.0f}\t{r['_our'] or '(空)'}\t{int(r['size'])}\t{r['sum']:,.0f}\t{sh:.0f}%\n")
    print(f"[{ent}] crosstab → {out.relative_to(ROOT)}")
    # divergence summary: their-labels whose biggest our-share < 90%
    print("  分歧大嘅 their_label（最大 our 佔比 <90%）:")
    for tl, sub in df.groupby("_their"):
        tot = sub["_amt"].sum()
        if tot <= 0 or tl == "": continue
        g = sub.groupby("_our")["_amt"].sum()
        top = g.max() / tot * 100
        if top < 90:
            mix = ", ".join(f"{k}:{v/tot*100:.0f}%" for k, v in g.sort_values(ascending=False).head(4).items() if v > 0)
            print(f"    {tl[:24]:24} {tot:>14,.0f}  → {mix}")

    if args.rows:
        rout = ROOT / "results" / f"{ent}_{yr_s}_{tag}_rows.tsv"
        pcol = fuzzy(df, cols.get("project", "")) or "project"
        adc = fuzzy(df, cols.get("account_desc", ""))
        keep = [c for c in (pcol, adc, "_their", "_our", "_amt") if c in df.columns or c.startswith("_")]
        sel = df[df["_their"] != df["_our"]][keep].copy()  # only where they differ (textually)
        sel.sort_values("_amt", ascending=False).head(3000).to_csv(rout, sep="\t", index=False)
        print(f"  row-level (textual mismatch, top 3000) → {rout.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
