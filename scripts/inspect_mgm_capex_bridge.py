"""MGM 23 non-gaming CAPEX — nail the Project Code → NG mapping (design the build fix).

inspect_mgm_capex_ng found CAPEX.xlsx 'JL details' carries several NG signals + a 'Confirmed to
include' sheet maps CER code (Project Code) → Section# (7.x=NGx / Gaming). This dumps, by $ coverage:
  (A) JL details: per NG-signal column, non-blank %, Σ|amt| covered, + full distinct values of the
      compact ones (項目性質 / KPMG-项目 / Non-gaming项目號 / Gaming-non-gaming / Section No.)
  (B) 'Confirmed to include': CER code → Section# + To Government Grouping (the authoritative map)
  (C) how much JL-details $ each signal can classify, alone and combined (best-available fallback)
so we pick the highest-coverage bridge and build the Section.1 assignment in build_mgm_23_raw.py.

Run (Windows):  python scripts/inspect_mgm_capex_bridge.py
Output: prints + results/inspect_mgm_capex_bridge.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
SIG_COLS = ["項目性質", "KPMG-项目", "Non-gaming项目號", "Gaming/non-gaming", "Section No.", "Section.1"]
AMT = "Debit minus Credit"
PCODE = "Project Code"


def find_file(name):
    hits = list((ROOT / "data").rglob(name))
    return hits[0] if hits else None


def main():
    L = ["# inspect_mgm_capex_bridge — design Project Code→NG for non-gaming CAPEX"]
    cap = find_file("CAPEX.xlsx")
    if not cap:
        L.append("X CAPEX.xlsx not found"); _w(L); return
    df = pd.read_excel(cap, sheet_name="JL details", header=0, dtype=object)
    a = pd.to_numeric(df[AMT], errors="coerce").fillna(0.0) if AMT in df.columns else pd.Series(0.0, index=df.index)
    tot = a.abs().sum() or 1
    L.append(f"\n## (A) JL details {len(df):,} rows  Σ|amt|={a.abs().sum():,.0f}")
    for c in SIG_COLS:
        if c not in df.columns:
            L.append(f"   {c:20s} — MISSING"); continue
        s = df[c].astype("string").fillna("").str.strip()
        cov = a.abs()[s.ne("")].sum()
        L.append(f"   {c:20s} nb{s.ne('').mean()*100:5.1f}%  $cov={cov/tot*100:5.1f}%  uniq={s[s.ne('')].nunique()}")
        if s[s.ne("")].nunique() <= 35:
            for v, x in a.abs()[s.ne("")].groupby(s[s.ne("")]).sum().sort_values(ascending=False).items():
                L.append(f"        {str(v)[:46]:46s} {x/1e6:8.2f}M")

    # (B) Confirmed to include — Project Code → Section# map
    try:
        ci = pd.read_excel(cap, sheet_name="Confirmed to include", header=0, dtype=object)
        cer = next((c for c in ci.columns if "CER" in str(c)), None)
        sec = next((c for c in ci.columns if "Section" in str(c)), None)
        gov = next((c for c in ci.columns if "Government" in str(c)), None)
        L.append(f"\n## (B) 'Confirmed to include' {len(ci):,} rows — CER='{cer}' Section#='{sec}' Gov='{gov}'")
        if sec:
            L.append("   Section# distinct: " + str(ci[sec].astype('string').value_counts().to_dict()))
        if cer and sec:
            L.append("   sample CER → Section# → Gov:")
            cols = [c for c in (cer, sec, gov) if c]
            for _, r in ci[cols].dropna(subset=[cer]).head(25).iterrows():
                L.append("      " + " | ".join(str(r[c])[:34] for c in cols))
    except Exception as e:
        L.append(f"\n## (B) Confirmed to include read failed: {e}")

    # (C) combined coverage: best-available fallback (項目性質 > KPMG-项目 > Non-gaming项目號 > Confirmed-join)
    have = pd.Series(False, index=df.index)
    for c in ("項目性質", "KPMG-项目", "Non-gaming项目號"):
        if c in df.columns:
            have |= df[c].astype("string").fillna("").str.strip().ne("")
    L.append(f"\n## (C) rows with ANY of (項目性質/KPMG-项目/Non-gaming项目號): "
             f"$cov={a.abs()[have].sum()/tot*100:.1f}%  ({int(have.sum()):,}/{len(df):,} rows)")
    if PCODE in df.columns:
        miss = ~have
        L.append(f"   rows with NONE — top Project Code (need Confirmed-join or manual):")
        pc = df[PCODE].astype("string").fillna("")
        for v, x in a.abs()[miss].groupby(pc[miss]).sum().sort_values(ascending=False).head(15).items():
            L.append(f"      {str(v)[:34]:34s} {x/1e6:8.2f}M")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_mgm_capex_bridge.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
