r"""Dump the DISTINCT projects that have BLANK vertical_label, per entity, from kpi_report.parquet.

This is the input to the targeted V re-tag: blank-V is localized (vml-23 / melco-23 / sjm), so
instead of re-LLM-ing ~6k subprojects we only classify the projects that are actually blank.

IMPORTANT — each entity stores its project/subproject under a DIFFERENT column name (vml
'SubProject_Name', melco 'Project name - Amended', sjm 'Project Name', mgm 'Project_code', …),
so a generic guess prints blanks even when the project IS populated. We use a per-entity column
list (mirrors each conf's columns.project + project_name_cols) AND print a diagnostic of every
project-ish column's non-blank count so the real holder is visible.

For blank-V rows: coalesce the entity's project cols → proj; group by (proj, subproject,
pt_class_V, ng_theme) → row count + Σ|amt|, sorted by Σ|amt| desc.

Run (Windows):  python scripts/dump_blank_v_projects.py              # all entities w/ blank V
                python scripts/dump_blank_v_projects.py sjm vml melco
Output: results/blank_v_<ent>.tsv  (≤3000 rows each, else _p1/_p2…) + console diagnostic.
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

# Per-entity project/subproject column candidates, in coalesce priority order.
# Mirrors each conf's `columns.project` + `project_name_cols` (the populated holders).
ENT_PROJ_COLS = {
    "galaxy": ["Project", "Sub Project", "Project Description", "Project Name", "項目名稱"],
    "sjm":    ["Project Name", "FileName"],
    "wynn":   ["Project", "Project Name", "Sub Project", "項目名稱"],
    "vml":    ["SubProject_Name", "Subproject", "項目名稱", "Project Name"],
    "melco":  ["Project name - Amended", "Project code - Amended", "initiative ID - Amended",
               "Project/Sub-Project Name", "项目名称&Unique ID", "initiative Name"],
    "mgm":    ["Project_code", "Project_name", "Project Name"],
}
AMT_CANDS = ["amount_mop", "Val/COArea Crcy", "MOP Amt", "Amount - Amended", "Debit minus Credit",
             "Reported Amount(MOP)", "Entry Voucher Amount/ Expense Amount", "adjusted_amount", "amount"]
PTV_CANDS = ["pt_class_V", "pt_class_v", "項目組V", "分類2", "投資領域", "分類1"]
NG_CANDS = ["項目類型", "ng_theme", "項目性質", "Section.1", "NG11 Category", "ng11_category"]
PER_CANDS = ["report_period", "report_bucket"]
PROJISH = ("project", "项目", "項目", "subproject", "sub project", "initiative", "proj")
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


def _coalesce(df, cols):
    out = pd.Series("", index=df.index, dtype="string")
    used = []
    for c in cols:
        if c not in df.columns:
            continue
        used.append(c)
        s = _s(df, c)
        fillmask = out.eq("") & s.ne("")
        out = out.mask(fillmask, s)
    return out, used


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
        if "vertical_label" not in df.columns:
            print(f"  [{alias}] no vertical_label col — skip"); continue
        blank = _s(df, "vertical_label").eq("")
        nb = int(blank.sum())
        if nb == 0:
            print(f"  [{alias}] blank V = 0 ✓")
            summary.append((alias, 0, 0, 0)); continue

        sub = df[blank].copy()
        # ── diagnostic: which project-ish columns actually hold a value in blank-V rows ──
        projish = [c for c in sub.columns if any(t in str(c).lower() for t in PROJISH)]
        diag = []
        for c in projish:
            n = int(_s(sub, c).ne("").sum())
            if n:
                diag.append((c, n))
        diag.sort(key=lambda kv: -kv[1])
        print(f"  [{alias}] blank-V rows={nb:,}")
        print(f"      project-ish cols (non-blank in blank-V rows): "
              + ("  ".join(f"{c!r}={n}" for c, n in diag[:10]) or "(NONE found!)"))

        amt_c = _col(sub, AMT_CANDS)
        amt = pd.to_numeric(sub[amt_c], errors="coerce").fillna(0.0).abs() if amt_c else pd.Series(0.0, index=sub.index)
        proj, used = _coalesce(sub, ENT_PROJ_COLS.get(alias, []))
        ptv_c, ng_c, per_c = _col(sub, PTV_CANDS), _col(sub, NG_CANDS), _col(sub, PER_CANDS)
        proj_filled = int(proj.ne("").sum())
        print(f"      coalesced project from {used} → {proj_filled:,}/{nb:,} rows have a project  "
              f"| pt_class_V col={ptv_c!r} ng col={ng_c!r} amt col={amt_c!r}")

        g = pd.DataFrame({
            "project":    proj,
            "pt_class_V": _s(sub, ptv_c),
            "ng_theme":   _s(sub, ng_c),
            "bucket":     _s(sub, per_c),
            "amt":        amt,
        })
        agg = (g.groupby(["project", "pt_class_V", "ng_theme"], dropna=False)
                 .agg(rows=("amt", "size"), buckets=("bucket", lambda s: ",".join(sorted(set(s)))),
                      amt_M=("amt", lambda s: round(s.sum() / 1e6, 2)))
                 .reset_index().sort_values("amt_M", ascending=False))
        summary.append((alias, nb, len(agg), round(amt.sum() / 1e6, 1)))

        outdir = ROOT / "results"; outdir.mkdir(exist_ok=True)
        if len(agg) <= BATCH:
            f = outdir / f"blank_v_{alias}.tsv"
            agg.to_csv(f, sep="\t", index=False)
            print(f"      → {f.relative_to(ROOT)}  ({len(agg):,} groups)")
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
