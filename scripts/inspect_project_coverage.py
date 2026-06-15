r"""Check whether the project / subproject dimension has any GAPS (blank rows) in each entity's
kpi_report.parquet — what feeds the Tableau `project` / `project_full` columns.

prep_tableau builds:
  project       = the entity's own project col (conf columns.project), trimmed
  subproject    = first matching subproject col
  project_full  = "project | subproject | description"  (display string for drill-down)

So if `project` is blank for some rows, `project_full` for those rows starts with " | …". This
script reports, per entity & per year_bucket: blank-project %, blank-subproject %, Σ|amt| of the
blank-project rows, and a few sample blank rows (their account/desc/vendor) so you can see WHAT is
missing a project.

Run (Windows):  python scripts/inspect_project_coverage.py
Output: prints + results/inspect_project_coverage.txt  (paste back)
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parent.parent
ENT = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
       "vml": "company_4", "melco": "company_5", "mgm": "company_6"}
SUB_CANDS = ["Sub project", "SubProject_Name", "Subproject_Name", "subproject", "项目名称中文",
             "项目英文名称", "Initiative Name", "Contents Name", "Subproject"]
AMT = ["amount_mop", "Val/COArea Crcy", "MOP Amt", "Amount - Amended", "Debit minus Credit",
       "Reported Amount(MOP)", "Entry Voucher Amount/ Expense Amount", "amount"]
_BLNK = ["", "nan", "None", "NaN", "<NA>", "NaT"]


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
    L = ["# inspect_project_coverage — blank project / subproject per entity & bucket"]
    for ent, com in ENT.items():
        L.append(f"\n{'='*72}\n## {ent}")
        p = ROOT / "data" / ent / "output" / f"{com}_kpi_report.parquet"
        if not p.exists():
            L.append(f"   (no kpi_report: {p.name})"); continue
        cfg = yaml.safe_load((ROOT / "conf" / com / "parameters.yml").read_text(encoding="utf-8"))
        proj_col = cfg.get("columns", {}).get("project", "")
        df = pd.read_parquet(p)
        n = len(df)
        rp = _s(df, _col(df, ["report_period", "year_bucket"])).replace("", "(none)")
        amt = pd.to_numeric(df[_col(df, AMT)], errors="coerce").fillna(0.0).abs() if _col(df, AMT) else pd.Series(0.0, index=df.index)

        pj = _s(df, proj_col if proj_col in df.columns else None)
        sub_col = _col(df, SUB_CANDS)
        sb = _s(df, sub_col)
        pj_blank = pj.isin(_BLNK)
        sb_blank = sb.isin(_BLNK)
        both_blank = pj_blank & sb_blank
        L.append(f"   rows={n:,}  project col={proj_col!r}  subproject col={sub_col!r}")
        L.append(f"   blank project   : {int(pj_blank.sum()):>8,} ({pj_blank.mean()*100:5.1f}%)  "
                 f"Σ|amt| {amt[pj_blank].sum()/1e6:,.1f}M")
        L.append(f"   blank subproject: {int(sb_blank.sum()):>8,} ({sb_blank.mean()*100:5.1f}%)")
        L.append(f"   blank BOTH      : {int(both_blank.sum()):>8,} ({both_blank.mean()*100:5.1f}%)  "
                 f"← these have NO project identity at all")
        if pj_blank.any():
            by = rp[pj_blank].value_counts()
            L.append("   blank-project by bucket: " + "  ".join(f"{k}={v}" for k, v in by.items()))
            # sample what the blank-project rows are
            ac = _col(df, ["account_code", "Account", "Cost Element", "Ledger Account", "ledger_account_id"])
            ad = _col(df, ["account_desc", "A/C Name", "Cost element descr.", "Ledger Hierarchy Level 5", "ledger_account"])
            vd = _col(df, ["vendor", "Vendor Short Name", "AC1 Vendor Name", "Supplier", "Vendor Name"])
            samp = df[pj_blank].head(6)
            L.append("   sample blank-project rows (acct | desc | vendor):")
            for _, r in samp.iterrows():
                L.append(f"      {str(r.get(ac,''))[:14]:14s} | {str(r.get(ad,''))[:30]:30s} | {str(r.get(vd,''))[:24]}")
    out = ROOT / "results" / "inspect_project_coverage.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste back")


if __name__ == "__main__":
    main()
