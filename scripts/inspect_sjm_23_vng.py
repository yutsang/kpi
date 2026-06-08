"""SJM 2023 V↔NG mismatch root-cause — PROVE where NG (項目性質) lives and why V disagrees.

User's point: the raw DEFINITELY has NG (項目性質) — it can't be missing. Correct. NG is present
and right (ng_label 吸引外國客源/文化藝術…). The BUG is V (vertical_id), derived from the project
NAME (step2 LLM + Project-Name keyword overrides), which disagrees with the project team's own
NG theme — giving nonsense like NG5 文化藝術 tagged V_THEME_PARK.

This dumps, for SJM 2023:
  (A) RAW sjm_23_raw.xlsx — 項目性質 distinct values + counts (proves the theme is in the raw).
  (B) tagged_rows — for every theme/NG/V column, non-blank% + distinct + top values
      (shows which surviving column still carries the theme → the one to drive V from).
  (C) ng_label × vertical_label crosstab (Σ|amt|, M) — visualises the mismatch.
  (D) 20 sample rows: project | 項目性質(raw) | ng_label | vertical_id — concrete mismatch.

Run (Windows):  python scripts/inspect_sjm_23_vng.py
Output: prints + results/inspect_sjm_23_vng.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
TR = ROOT / "data" / "sjm" / "interim" / "company_2_tagged_rows.parquet"
AMT = "Val/COArea Crcy"
THEME_COLS = ["項目性質", "ng_code", "ng_label", "ng11_category", "項目類型",
              "vertical_id", "vertical_label", "horizontal_label"]


def _raw_theme(L):
    """(A) Read the built raw and show 項目性質 distinct — prove NG is in the source."""
    cands = list((ROOT / "data").rglob("sjm_23_raw.xlsx"))
    if not cands:
        L.append("\n## (A) RAW sjm_23_raw.xlsx NOT FOUND under data/**"); return
    fp = cands[0]
    try:
        raw = pd.read_excel(fp, sheet_name="combine", dtype=object)
    except Exception as e:
        L.append(f"\n## (A) RAW read failed: {e}"); return
    L.append(f"\n## (A) RAW {fp.relative_to(ROOT)}  rows={len(raw):,}  cols={list(raw.columns)}")
    if "項目性質" in raw.columns:
        s = raw["項目性質"].astype("string").fillna("").str.strip()
        L.append(f"   項目性質 non-blank={s.ne('').mean()*100:.1f}%  distinct={s[s.ne('')].nunique()}")
        for v, n in s[s.ne("")].value_counts().items():
            L.append(f"      {str(v)[:30]:30s} {n:,}")
    else:
        L.append("   !! 項目性質 NOT a column in the raw")


def main():
    L = ["# inspect_sjm_23_vng — prove NG location + V↔NG mismatch"]
    _raw_theme(L)

    if not TR.exists():
        L.append(f"\nX tagged_rows {TR} missing — run kedro sjm first"); _w(L); return
    df = pd.read_parquet(TR)
    per = next((c for c in ("report_period", "report_year", "years") if c in df.columns), None)
    if per:
        df = df[df[per].astype(str).str.startswith("23")].copy()
    amt = AMT if AMT in df.columns else next((c for c in df.columns if "Crcy" in str(c) or "Amount" in str(c)), None)
    a = pd.to_numeric(df[amt], errors="coerce").fillna(0.0)
    tot = a.abs().sum() or 1
    L.append(f"\n## (B) tagged_rows 23 rows={len(df):,}  Σ={a.sum():,.0f}  amount={amt!r}")
    L.append(f"   ALL columns: {list(df.columns)}")
    for c in THEME_COLS:
        if c not in df.columns:
            L.append(f"   {c:14s} — MISSING"); continue
        s = df[c].astype("string").fillna("").str.strip()
        nb = s.ne("").mean() * 100
        top = " | ".join(f"{v}({n})" for v, n in s[s.ne("")].value_counts().head(6).items())
        L.append(f"   {c:14s} nb{nb:5.1f}%  uniq{s[s.ne('')].nunique():>4}  {top[:110]}")

    # (C) ng_label × vertical_label crosstab
    ngc = next((c for c in ("ng_label", "ng_code") if c in df.columns), None)
    if ngc and "vertical_label" in df.columns:
        L.append(f"\n## (C) {ngc} × vertical_label  (Σ|amt| M) — off-diagonal = mismatch:")
        ct = pd.crosstab(df[ngc].astype("string").fillna("(blank)"),
                         df["vertical_label"].astype("string").fillna("(blank)"),
                         values=a.abs(), aggfunc="sum").fillna(0) / 1e6
        L.append(ct.round(1).to_string())

    # (D) 20 sample mismatched rows
    proj = next((c for c in ("Project Name", "project_name", "項目名稱") if c in df.columns), None)
    cols = [c for c in (proj, "項目性質", ngc, "vertical_id", "vertical_label") if c and c in df.columns]
    if cols:
        L.append(f"\n## (D) 20 sample rows  [{', '.join(cols)}]:")
        for _, r in df[cols].head(20).iterrows():
            L.append("   " + " | ".join(str(r[c])[:24] for c in cols))
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_sjm_23_vng.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
