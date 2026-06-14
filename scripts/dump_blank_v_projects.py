r"""Dump the DISTINCT projects that have BLANK vertical_label, per entity, from kpi_report.parquet.

This is the input to the targeted V re-tag: blank-V is localized (vml-23 / melco-23 / sjm all-bucket),
so instead of re-LLM-ing ~6k subprojects we only need to classify the handful of projects that are
actually blank — using the PROJECT TEAM's own V column (pt_class_V) as the ground truth, mapped to
our 26-V taxonomy.

For each entity, among rows where vertical_label is blank:
  group by (project, subproject, pt_class_V, ng_theme) → row count + Σ|amt|, sorted by Σ|amt| desc.

Run (Windows):  python scripts/dump_blank_v_projects.py              # all entities w/ blank V
                python scripts/dump_blank_v_projects.py sjm vml melco
Output: results/blank_v_<ent>.tsv  (one per entity; ≤3000 rows each, else split _p1/_p2…)
        + a short console summary. Paste the TSVs back.
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
ENT = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
       "vml": "company_4", "melco": "company_5", "mgm": "company_6"}
AMT_CANDS = ["amount_mop", "Val/COArea Crcy", "MOP Amt", "Amount - Amended", "Debit minus Credit",
             "Reported Amount(MOP)", "Entry Voucher Amount/ Expense Amount", "amount"]
PROJ_CANDS = ["Project Name", "Project", "project", "項目名稱", "投資項目名稱",
              "Project Description", "project_name", "Project Name "]
SUB_CANDS = ["subproject", "子項目", "Sub Project", "SubProject", "細項", "sub_project", "subproject_name"]
PTV_CANDS = ["pt_class_V", "pt_class_v", "項目組V", "分類1"]
NG_CANDS = ["項目類型", "ng_theme", "NG11 Category", "項目性質", "ng11_category"]
PER_CANDS = ["report_period", "report_bucket"]
BATCH = 3000


def _col(df, cands):
    for c in cands:
        if c in df.columns:
            return c
    low = {str(c).lower(): c for c in df.columns}
    for c in cands:
        if str(c).lower() in low:
            return low[str(c).lower()]
    return None


def _s(df, c):
    if c is None or c not in df.columns:
        return pd.Series("", index=df.index, dtype="string")
    return df[c].astype("string").fillna("").str.strip()


def main():
    want = [a.lower() for a in sys.argv[1:]] or list(ENT)
    summary = []
    for alias in want:
        com = ENT.get(alias)
        if not com:
            print(f"  ?? unknown entity {alias!r}"); continue
        p = ROOT / "data" / alias / "output" / f"{com}_kpi_report.parquet"
        if not p.exists():
            print(f"  [{alias}] no kpi_report.parquet — skip"); continue
        df = pd.read_parquet(p)
        vcol = "vertical_label"
        if vcol not in df.columns:
            print(f"  [{alias}] no vertical_label col — skip"); continue
        blank = _s(df, vcol).eq("")
        nb = int(blank.sum())
        if nb == 0:
            print(f"  [{alias}] blank V = 0 ✓  (nothing to dump)")
            summary.append((alias, 0, 0, 0)); continue

        sub = df[blank].copy()
        amt_c = _col(sub, AMT_CANDS)
        amt = pd.to_numeric(sub[amt_c], errors="coerce").fillna(0.0).abs() if amt_c else pd.Series(0.0, index=sub.index)
        proj_c, subp_c, ptv_c, ng_c, per_c = (_col(sub, PROJ_CANDS), _col(sub, SUB_CANDS),
                                              _col(sub, PTV_CANDS), _col(sub, NG_CANDS), _col(sub, PER_CANDS))
        g = pd.DataFrame({
            "project":   _s(sub, proj_c),
            "subproject": _s(sub, subp_c),
            "pt_class_V": _s(sub, ptv_c),
            "ng_theme":  _s(sub, ng_c),
            "bucket":    _s(sub, per_c),
            "amt":       amt,
        })
        agg = (g.groupby(["project", "subproject", "pt_class_V", "ng_theme"], dropna=False)
                 .agg(rows=("amt", "size"), buckets=("bucket", lambda s: ",".join(sorted(set(s)))),
                      amt_M=("amt", lambda s: round(s.sum() / 1e6, 2)))
                 .reset_index().sort_values("amt_M", ascending=False))
        ptv_fill = int(g["pt_class_V"].ne("").sum())
        print(f"  [{alias}] blank-V rows={nb:,}  distinct project-groups={len(agg):,}  "
              f"Σ|amt|={amt.sum()/1e6:,.1f}M  | cols: proj={proj_c!r} pt_class_V={ptv_c!r} "
              f"({ptv_fill:,}/{nb:,} rows have pt_class_V) ng={ng_c!r}")
        summary.append((alias, nb, len(agg), round(amt.sum() / 1e6, 1)))

        outdir = ROOT / "results"; outdir.mkdir(exist_ok=True)
        if len(agg) <= BATCH:
            f = outdir / f"blank_v_{alias}.tsv"
            agg.to_csv(f, sep="\t", index=False)
            print(f"      → {f.relative_to(ROOT)}")
        else:
            for i in range(0, len(agg), BATCH):
                f = outdir / f"blank_v_{alias}_p{i // BATCH + 1}.tsv"
                agg.iloc[i:i + BATCH].to_csv(f, sep="\t", index=False)
                print(f"      → {f.relative_to(ROOT)}  ({i + 1}-{min(i + BATCH, len(agg))})")

    print("\n# summary  entity | blank_V_rows | distinct_groups | Σ|amt|M")
    for a, r, d, m in summary:
        print(f"   {a:7s} {r:>9,} {d:>9,} {m:>12,.1f}")


if __name__ == "__main__":
    main()
