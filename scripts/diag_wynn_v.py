"""Diagnose wynn (company_3) vertical_id = blank for ~100% of rows, while horizontal works.
ALSO validates the new row_vertical_overrides column_map on '項目分類2' (THEME2V): confirms the
exact category-column NAME + that its distinct VALUES are byte-exact map keys, per report_period.

Run (Windows):  python scripts/diag_wynn_v.py
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
# the column_map I wired in conf/company_3 row_vertical_overrides — confirm name + values:
CATCOLS = ["項目分類2", "項目分類1", "範疇第二層標籤", "項目性質", "comp费用大类", "Annex 2 Summary Cateogry"]
THEME2V = {
    "博彩娛樂場場地的優化": "V_GAMING_VENUE", "博彩設施及設備的優化": "V_GAMING_EQUIP",
    "吸引外國客源": "V_INVITE_GUEST", "會議展覽": "V_MICE", "娛樂表演": "V_CONCERT",
    "體育盛事": "V_SPORT_EVENT", "文化藝術": "V_ART_EXHIBITION", "健康養生": "V_WELLNESS",
    "主題遊樂": "V_THEME_PARK", "美食之都": "V_RESTAURANT", "社區旅遊": "V_COMMUNITY",
    "海上旅遊": "V_MARITIME", "其他": "V_OTHER",
}


def main():
    L = ["# diag_wynn_v"]

    # ── 1. unique_projects.xlsx — is V tagged there? ──
    up = INTERIM / "company_3_unique_projects.xlsx"
    if up.exists():
        p = pd.read_excel(up)
        L.append(f"\n## unique_projects.xlsx  ({len(p):,} project rows)")
        for c in ("llm_vertical", "manual_vertical", "vertical_id"):
            if c in p.columns:
                nb = p[c].astype("string").fillna("").str.strip().ne("").sum()
                L.append(f"   {c:16s}: {nb:,}/{len(p):,} non-blank ({nb/max(len(p),1)*100:.1f}%)")
    else:
        L.append(f"\n## unique_projects.xlsx  X NOT FOUND ({up})")

    # ── 2. tagged_rows.parquet — V hole + category-column coverage ──
    tr = INTERIM / "company_3_tagged_rows.parquet"
    if not tr.exists():
        L.append(f"\n## tagged_rows.parquet  X NOT FOUND ({tr})")
        _write(L); return
    df = pd.read_parquet(tr)
    amt = AMT if AMT in df.columns else next((c for c in df.columns if "Entry Voucher" in str(c)), None)
    a = pd.to_numeric(df[amt], errors="coerce").fillna(0.0) if amt else pd.Series(0.0, index=df.index)
    atot = a.abs().sum() or 1
    per = next((c for c in ("report_period", "report_year", "years") if c in df.columns), None)
    L.append(f"\n## tagged_rows.parquet  ({len(df):,} rows, Σ={a.sum():,.0f})  amount={amt!r}")
    L.append(f"   ALL columns: {list(df.columns)}")

    if PROJ_COL in df.columns:
        pj = df[PROJ_COL].astype("string").fillna("").str.strip()
        L.append(f"   '{PROJ_COL}' blank: {pj.eq('').sum():,} rows / {a.abs()[pj.eq('')].sum()/atot*100:.1f}% of |amt|")
    if "vertical_id" in df.columns:
        vid = df["vertical_id"].astype("string").fillna("(blank)").replace({"": "(blank)", "nan": "(blank)"})
        L.append("   vertical_id by |amount|:")
        for v, s in a.abs().groupby(vid).sum().sort_values(ascending=False).head(20).items():
            L.append(f"      {str(v)[:26]:26s} {s/atot*100:6.1f}%")

    # category columns: which exist + value_counts + THEME2V membership ✓/✗
    L.append("\n## V-category column coverage (validates the 項目分類2 column_map)")
    for col in CATCOLS:
        if col not in df.columns:
            L.append(f"   [{col}] — NOT a column"); continue
        s = df[col].astype("string").fillna("").str.strip()
        blankpct = a.abs()[s.eq("")].sum() / atot * 100
        L.append(f"   [{col}] present — blank {blankpct:.1f}% of |amt|; values (by |amt|, ✓=in THEME2V):")
        for v, amt_s in a.abs().groupby(s).sum().sort_values(ascending=False).head(25).items():
            if v == "": continue
            mark = "✓" if v in THEME2V else "✗ NOT-IN-MAP"
            L.append(f"      {mark:14s} {str(v)[:30]:30s} {amt_s/atot*100:6.1f}%")
        # per-period presence (does 24/23 carry this column with values?)
        if per:
            seg = []
            for pp, g in df.groupby(df[per].astype(str)):
                gs = g[col].astype("string").fillna("").str.strip()
                seg.append(f"{pp}:{gs.ne('').mean()*100:.0f}%nonblank")
            L.append(f"      per-period nonblank: {'  '.join(seg)}")

    _write(L)


def _write(L):
    out = ROOT / "results" / "diag_wynn_v.txt"
    out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
