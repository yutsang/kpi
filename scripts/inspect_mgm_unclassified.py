"""MGM — why 14,973 rows are NG 未分類 (Section.1 → no NGn).

step6 logs '[mgm] NG from databook col(s) ['Section.1']: 230,184/245,157 mapped (14,973 → 未分類)'.
NG comes ONLY from Section.1 (the project-team 範疇 col). 14,973 rows resolve to no NGn — either
Section.1 is blank, or its value doesn't normalize to NG0-11. This finds the pattern: which Source /
year / Project_code / Section.1 raw value those rows have, so we can fix the source (build/section
inference) rather than guess.

Run (Windows):  python scripts/inspect_mgm_unclassified.py
Output: prints + results/inspect_mgm_unclassified.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
TR = ROOT / "data" / "mgm" / "interim" / "company_6_tagged_rows.parquet"
AMT = "Debit minus Credit"
NGCOL = "Section.1"


def main():
    L = ["# inspect_mgm_unclassified — why Section.1 → 未分類 (no NGn)"]
    if not TR.exists():
        L.append(f"X {TR} missing"); _w(L); return
    df = pd.read_parquet(TR)
    a = pd.to_numeric(df[AMT], errors="coerce").fillna(0.0)
    sys.path.insert(0, str(ROOT / "src"))
    from kpi.lib.conf import load_categories
    from kpi.pipelines.step2_tag_projects._logic import normalize_ng_code
    cats = load_categories()
    sec = df[NGCOL].astype("string").fillna("") if NGCOL in df.columns else pd.Series("", index=df.index)
    ngc = sec.map({s: (normalize_ng_code(s, cats) or "") for s in sec.unique()})
    bad = ~ngc.str.fullmatch(r"NG\d+").fillna(False)
    L.append(f"\nTOTAL rows={len(df):,}  未分類(no NGn)={int(bad.sum()):,}  Σ|amt|={a.abs()[bad].sum():,.0f}")

    def dist(col, title, n=15):
        if col not in df.columns:
            L.append(f"\n## {title}: col {col!r} MISSING"); return
        s = df[col].astype("string").fillna("(blank)")
        L.append(f"\n## 未分類 by {title} ({col}) — Σ|amt| M / rows:")
        g = pd.DataFrame({"k": s[bad], "amt": a.abs()[bad]}).groupby("k").agg(amt=("amt", "sum"), n=("amt", "size")).sort_values("amt", ascending=False)
        for k, row in g.head(n).iterrows():
            L.append(f"   {str(k)[:38]:38s} {row['amt']/1e6:9.2f}M  ({int(row['n'])} rows)")

    dist("report_period", "year")
    dist("Source", "Source")
    dist(NGCOL, "Section.1 raw value")
    dist("Section", "Section")
    dist("Project_code", "Project_code")
    dist("Project_name", "Project_name")
    dist("vertical_label", "vertical_label (what V they got)")
    dist("Ledger Account", "Ledger Account")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_mgm_unclassified.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
