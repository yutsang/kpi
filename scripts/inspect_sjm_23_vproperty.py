"""SJM 2023 — ROOT CAUSE of V_PROPERTY_UPGRADE landing on 海上旅遊/會議展覽 rows.

NOT a fallback-removal question — find WHY. SJM V is per-project broadcast (one project → one V,
applied to all its rows). So a project step2 tagged V_PROPERTY_UPGRADE will stamp that on every
row, even rows whose theme (項目類型) is 海上旅遊 (NG10) / 會議展覽 (NG2).

This dumps, for SJM 2023 rows where vertical_id == V_PROPERTY_UPGRADE:
  (A) the distinct projects (Project Name / Project Code) + their theme(s) + vertical_source + Σ|amt|
  (B) for EACH such project, its FULL theme distribution across all its 23 rows — detects whether the
      project is genuinely mono-theme 海上旅遊 (LLM mis-pick) or multi-theme (broadcast smear)
  (C) what step2 actually assigned this project in company_2_unique_projects.xlsx (llm_vertical +
      any source/reason cols) — so we see if the LLM picked V_PROPERTY_UPGRADE and on what basis

Run (Windows):  python scripts/inspect_sjm_23_vproperty.py
Output: prints + results/inspect_sjm_23_vproperty.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
TR = ROOT / "data" / "sjm" / "interim" / "company_2_tagged_rows.parquet"
UP = ROOT / "data" / "sjm" / "interim" / "company_2_unique_projects.xlsx"
AMT = "Val/COArea Crcy"
PROJ = "Project Name"
THEME = "項目類型"
TARGET = "V_PROPERTY_UPGRADE"


def main():
    L = [f"# inspect_sjm_23_vproperty — root cause of {TARGET} on 海上旅遊/會議展覽"]
    if not TR.exists():
        L.append(f"X {TR} missing"); _w(L); return
    df = pd.read_parquet(TR)
    per = next((c for c in ("report_period", "report_year", "years") if c in df.columns), None)
    if per:
        df = df[df[per].astype(str).str.startswith("23")].copy()
    a = pd.to_numeric(df[AMT], errors="coerce").fillna(0.0)
    vid = df["vertical_id"].astype("string").fillna("")
    proj = df[PROJ].astype("string").fillna("(blank)") if PROJ in df.columns else pd.Series("(no col)", index=df.index)
    theme = df[THEME].astype("string").fillna("(blank)") if THEME in df.columns else pd.Series("(no col)", index=df.index)
    vsrc = df["vertical_source"].astype("string").fillna("") if "vertical_source" in df.columns else pd.Series("", index=df.index)

    m = vid.eq(TARGET)
    L.append(f"\n{TARGET} rows={int(m.sum()):,}  Σ|amt|={a.abs()[m].sum():,.0f}")

    # (A) distinct (project, theme, source)
    L.append(f"\n## (A) {TARGET} rows by (Project Name, 項目類型, vertical_source):")
    g = (pd.DataFrame({"proj": proj[m], "theme": theme[m], "src": vsrc[m], "amt": a.abs()[m]})
         .groupby(["proj", "theme", "src"]).agg(n=("amt", "size"), amt=("amt", "sum"))
         .sort_values("amt", ascending=False))
    for (p, t, s), row in g.iterrows():
        L.append(f"   {str(p)[:34]:34s} | {str(t)[:10]:10s} | {str(s)[:10]:10s} | {int(row['n']):>4} rows | {row['amt']/1e6:7.2f}M")

    # (B) for each such project, its FULL theme spread across all its 23 rows
    bad_projs = list(dict.fromkeys(proj[m].tolist()))
    L.append(f"\n## (B) for each {TARGET} project — its FULL theme(項目類型) spread (all its 23 rows):")
    for p in bad_projs[:20]:
        pm = proj.eq(p)
        spread = a.abs()[pm].groupby(theme[pm]).sum().sort_values(ascending=False)
        vspread = vid[pm].value_counts().to_dict()
        L.append(f"   [{str(p)[:40]}]  rows={int(pm.sum())}  V={vspread}")
        for t, x in spread.items():
            L.append(f"        theme={str(t)[:14]:14s} {x/1e6:7.2f}M")

    # (C) what step2 gave these projects
    if UP.exists():
        up = pd.read_excel(UP)
        pcol = next((c for c in up.columns if "roject" in str(c) or c == PROJ), up.columns[0])
        L.append(f"\n## (C) company_2_unique_projects rows for these projects (cols: {list(up.columns)[:12]}):")
        sel = up[up[pcol].astype("string").isin([str(x) for x in bad_projs])]
        show = [c for c in (pcol, "manual_vertical", "llm_vertical", "tag_source", "vertical_reason", "reason", "llm_reason") if c in up.columns]
        for _, r in sel[show].head(30).iterrows():
            L.append("   " + " | ".join(f"{c}={str(r[c])[:40]}" for c in show))
    else:
        L.append(f"\n## (C) {UP.name} not found")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_sjm_23_vproperty.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
