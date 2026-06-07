"""Diagnose wynn (company_3) vertical_id = blank for ~100% of rows in every period, while
horizontal works. Three candidate causes — this dump decides which:
  (a) unique_projects.xlsx llm_vertical never populated (step2 never tagged wynn)
  (b) most rows have a BLANK project col ('Name of Investment Project') -> nothing to join
  (c) join-key mismatch (project values in rows not present as unique_projects keys)

Run (Windows):
  python scripts/diag_wynn_v.py
Output: prints + results/diag_wynn_v.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
INTERIM = ROOT / "data" / "wynn" / "interim"
PROJ_COL = "Name of Investment Project"
AMT = "Entry Voucher Amount/ Expense Amount "   # trailing space is real
ALT_PROJ = ["Annex 2 Summary Cateogry", "Sub project", "项目名称中文", "項目性質"]


def main():
    L = ["# diag_wynn_v"]

    # ── 1. unique_projects.xlsx — is V actually tagged? ──────────────────────
    up = INTERIM / "company_3_unique_projects.xlsx"
    if up.exists():
        p = pd.read_excel(up)
        L.append(f"\n## unique_projects.xlsx  ({len(p):,} project rows)")
        L.append(f"   columns: {list(p.columns)}")
        for c in ("llm_vertical", "manual_vertical", "vertical_id"):
            if c in p.columns:
                nonblank = p[c].astype("string").fillna("").str.strip().ne("").sum()
                L.append(f"   {c:16s}: {nonblank:,}/{len(p):,} non-blank ({nonblank/max(len(p),1)*100:.1f}%)")
        vcol = next((c for c in ("manual_vertical", "llm_vertical", "vertical_id") if c in p.columns), None)
        if vcol:
            L.append(f"   top {vcol} values:")
            for v, n in p[vcol].astype("string").fillna("(blank)").value_counts().head(15).items():
                L.append(f"      {str(v)[:40]:40s} {n:,}")
        if PROJ_COL in p.columns:
            L.append(f"   sample {PROJ_COL} keys: " +
                     " | ".join(map(str, p[PROJ_COL].dropna().astype(str).head(8))))
    else:
        L.append(f"\n## unique_projects.xlsx  X NOT FOUND at {up}")

    # ── 2. tagged_rows.parquet — where does V go blank? ──────────────────────
    tr = INTERIM / "company_3_tagged_rows.parquet"
    if tr.exists():
        df = pd.read_parquet(tr)
        amt = AMT if AMT in df.columns else next((c for c in df.columns if "Entry Voucher" in str(c)), None)
        a = pd.to_numeric(df[amt], errors="coerce").fillna(0.0) if amt else pd.Series(0.0, index=df.index)
        atot = a.abs().sum() or 1
        L.append(f"\n## tagged_rows.parquet  ({len(df):,} rows, Σ={a.sum():,.0f})  amount_col={amt!r}")
        L.append(f"   columns: {list(df.columns)}")

        # project col present & populated?
        if PROJ_COL in df.columns:
            pj = df[PROJ_COL].astype("string").fillna("").str.strip()
            blank_proj_amt = a.abs()[pj.eq("")].sum()
            L.append(f"   '{PROJ_COL}' blank: {pj.eq('').sum():,} rows / {blank_proj_amt/atot*100:.1f}% of |amount|")
            L.append(f"   '{PROJ_COL}' distinct non-blank: {pj[pj.ne('')].nunique():,}")
            L.append(f"   sample row project values: " + " | ".join(map(str, pj[pj.ne('')].head(8))))
        else:
            L.append(f"   X '{PROJ_COL}' NOT a column in tagged_rows — V can never join. cols above ^")

        # vertical_id distribution (by |amount|)
        if "vertical_id" in df.columns:
            vid = df["vertical_id"].astype("string").fillna("(blank)")
            L.append("   vertical_id by |amount|:")
            g = a.abs().groupby(vid).sum().sort_values(ascending=False)
            for v, s in g.head(20).items():
                L.append(f"      {str(v)[:24]:24s} {s/atot*100:6.1f}%  ({s:,.0f})")

            # for BLANK-V rows: what's in the project col + alternatives?
            blankv = vid.isin(["(blank)", "", "nan", "None"])
            L.append(f"\n   --- BLANK vertical_id rows = {a.abs()[blankv].sum()/atot*100:.1f}% of |amount| ---")
            if PROJ_COL in df.columns:
                sub = df.loc[blankv, PROJ_COL].astype("string").fillna("(blank)")
                suba = a.abs()[blankv]
                L.append(f"   their '{PROJ_COL}' top values (by |amount|):")
                for v, s in suba.groupby(sub).sum().sort_values(ascending=False).head(15).items():
                    L.append(f"      {str(v)[:40]:40s} {s/atot*100:6.1f}%")
            for ac in ALT_PROJ:
                if ac in df.columns:
                    sub = df.loc[blankv, ac].astype("string").fillna("(blank)")
                    suba = a.abs()[blankv]
                    L.append(f"   blank-V rows '{ac}' top values (by |amount|):")
                    for v, s in suba.groupby(sub).sum().sort_values(ascending=False).head(10).items():
                        L.append(f"      {str(v)[:40]:40s} {s/atot*100:6.1f}%")
        else:
            L.append("   X no vertical_id column in tagged_rows")
    else:
        L.append(f"\n## tagged_rows.parquet  X NOT FOUND at {tr}")

    out = ROOT / "results" / "diag_wynn_v.txt"
    out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
