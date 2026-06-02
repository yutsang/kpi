"""Per-project view: project-team raw NG (項目性質) vs OUR vertical/NG + amount — for one
entity+year. Also dumps BLANK-project rows in full so we can see where stray totals come from.

Use cases:
  · 邊個 project 我哋 V 同項目組 項目性質 唔夾（寫 targeted override）
  · 搵某範疇嘅 project code（e.g. Wynn 24 健康養生 用咩 code）
  · 查 blank-project 大額行係邊嚟（e.g. SJM 24 嗰 485M 2 行）

Run (Windows):
  python scripts/inspect_proj_ng.py --entity wynn --year 24 --grep 健康
  python scripts/inspect_proj_ng.py --entity sjm  --year 24 --blanks
  python scripts/inspect_proj_ng.py --entity vml  --year 24 --top 60
Output: results/<ent>_<year>_projng.tsv
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
ENTITIES = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
            "vml": "company_4", "melco": "company_5", "mgm": "company_6"}
YEAR_CANDIDATES = ("report_period", "report_year", "Yr related", "years")


def fuzzy(df, name):
    if not name: return None
    if name in df.columns: return name
    for c in df.columns:
        if str(c).strip() == str(name).strip(): return c
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", required=True, choices=sorted(ENTITIES))
    ap.add_argument("--year", required=True)
    ap.add_argument("--top", type=int, default=50)
    ap.add_argument("--grep", default="", help="只列 project 名/項目性質 含呢個字嘅")
    ap.add_argument("--blanks", action="store_true", help="淨係 dump blank-project 行嘅全欄")
    args = ap.parse_args()
    ent, com, tag = args.entity, ENTITIES[args.entity], str(args.year)

    parquet = ROOT / "data" / ent / "output" / f"{com}_kpi_report.parquet"
    if not parquet.exists():
        print(f"X missing {parquet.relative_to(ROOT)}"); return
    cf = yaml.safe_load((ROOT / "conf" / com / "parameters.yml").open(encoding="utf-8"))
    cols = cf.get("columns", {}) or {}
    df = pq.read_table(parquet).replace_schema_metadata(None).to_pandas()
    ycol = next((c for c in YEAR_CANDIDATES if c in df.columns), None)
    yr = df[ycol].astype("string").fillna("")
    df = df[yr.str.startswith(tag) | (yr == f"Yr 20{tag}") | (yr == f"20{tag}")].copy()

    amt = fuzzy(df, cols.get("amount", ""))
    df["_amt"] = pd.to_numeric(df[amt], errors="coerce").fillna(0) if amt else 0.0
    pcol = fuzzy(df, cols.get("project", "")) or "project"
    ngc = fuzzy(df, cols.get("ng11_category", ""))     # 項目性質 / NG11 Category
    vlab = "vertical_label" if "vertical_label" in df.columns else "vertical_id"
    pstr = df[pcol].astype("string").fillna("").str.strip()

    out = ROOT / "results" / f"{ent}_{tag}_projng.tsv"
    out.parent.mkdir(parents=True, exist_ok=True)

    if args.blanks:
        blank = df[pstr == ""]
        keep = [c for c in (cols.get("unique_id"), cols.get("account_code"), cols.get("account_desc"),
                            cols.get("description"), ngc, ycol) if c and c in df.columns]
        with out.open("w", encoding="utf-8") as f:
            f.write(f"# {ent} {tag} BLANK-project rows: {len(blank)} rows, {blank['_amt'].sum():,.0f}\n")
            f.write("amount\t" + "\t".join(str(k) for k in keep) + "\n")
            for _, r in blank.sort_values("_amt", ascending=False).head(args.top).iterrows():
                f.write(f"{r['_amt']:,.0f}\t" + "\t".join(str(r[k])[:40] for k in keep) + "\n")
        print(f"[{ent} {tag}] {len(blank)} blank-project rows, {blank['_amt'].sum():,.0f} → {out.relative_to(ROOT)}")
        for _, r in blank.sort_values("_amt", ascending=False).head(12).iterrows():
            print(f"   {r['_amt']:>16,.0f} | " + " | ".join(str(r[k])[:26] for k in keep))
        return

    rows = []
    for proj, sub in df.groupby(pstr, dropna=False):
        their = ""
        if ngc:
            vc = sub[ngc].astype(str).str.strip().replace({"nan": ""})
            vc = vc[vc != ""]
            their = vc.value_counts().idxmax() if len(vc) else ""
        ov = sub[vlab].astype(str)
        ov = ov[ov != "nan"]
        our = ov.value_counts().idxmax() if len(ov) else ""
        rows.append((str(proj), their, our, len(sub), float(sub["_amt"].sum())))
    g = pd.DataFrame(rows, columns=["project", "their_項目性質", "our_V", "n_rows", "amount"])
    if args.grep:
        m = g["project"].str.contains(args.grep, case=False, na=False) | \
            g["their_項目性質"].str.contains(args.grep, case=False, na=False)
        g = g[m]
    g = g.sort_values("amount", ascending=False).head(args.top)
    with out.open("w", encoding="utf-8") as f:
        f.write(f"# {ent} {tag}  項目組欄={ngc!r}  (their_項目性質 vs our_V)\n")
        f.write("project\ttheir_項目性質\tour_V\tn_rows\tamount\n")
        for _, r in g.iterrows():
            f.write(f"{str(r['project'])[:60]}\t{str(r['their_項目性質'])[:24]}\t{r['our_V']}\t{int(r['n_rows'])}\t{r['amount']:,.0f}\n")
    print(f"[{ent} {tag}] {len(g)} projects → {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
