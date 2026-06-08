"""SJM V check after the 澳門皇宮→V_MARITIME / 多功能廳→V_MICE overrides (+ eligibility).

The 2 keyword overrides are NOT year-gated, so they hit 23 AND 24/25. This verifies:
  (1) per report_period, V_PROPERTY_UPGRADE remaining — project / theme / count / Σ|amt|
      (SJM 23 should be ~0 now; see whether 24/25 still carry legit ones)
  (2) every row matched by '澳門皇宮' or '多功能廳' — period / project / theme(項目類型) / V — to
      confirm the override is theme-consistent everywhere (not over-reaching in 24/25)
  (3) per report_period out-of-bucket count (V not in its NG eligible_verticals) — eligibility only
      gates "23", so this shows whether 24/25 have residual V↔NG mismatches worth gating later

Run (Windows):  python scripts/inspect_sjm_v_check.py
Output: prints + results/inspect_sjm_v_check.txt
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
PROJ = "Project Name"
THEME = "項目類型"
KEYWORDS = ["澳門皇宮", "多功能廳"]


def main():
    L = ["# inspect_sjm_v_check — V_PROPERTY_UPGRADE + 澳門皇宮/多功能廳 override audit (all years)"]
    if not TR.exists():
        L.append(f"X {TR} missing"); _w(L); return
    df = pd.read_parquet(TR)
    a = pd.to_numeric(df[AMT], errors="coerce").fillna(0.0)
    rp = df["report_period"].astype("string").fillna("(blank)")
    proj = df[PROJ].astype("string").fillna("") if PROJ in df.columns else pd.Series("", index=df.index)
    theme = df[THEME].astype("string").fillna("") if THEME in df.columns else pd.Series("", index=df.index)
    vid = df["vertical_id"].astype("string").fillna("")
    vlab = df["vertical_label"].astype("string").fillna("") if "vertical_label" in df.columns else vid

    # (1) V_PROPERTY_UPGRADE remaining per period
    m = vid.eq("V_PROPERTY_UPGRADE")
    L.append(f"\n## (1) V_PROPERTY_UPGRADE remaining: {int(m.sum()):,} rows  Σ|amt|={a.abs()[m].sum():,.0f}")
    if m.any():
        g = (pd.DataFrame({"period": rp[m], "proj": proj[m], "theme": theme[m], "amt": a.abs()[m]})
             .groupby(["period", "proj", "theme"]).agg(n=("amt", "size"), amt=("amt", "sum"))
             .sort_values("amt", ascending=False))
        for (p, pr, t), row in g.head(30).iterrows():
            L.append(f"   {str(p):9s} | {str(pr)[:30]:30s} | {str(t)[:10]:10s} | {int(row['n']):>4} | {row['amt']/1e6:7.2f}M")

    # (2) rows matched by the 2 keywords
    for kw in KEYWORDS:
        mm = proj.str.contains(kw, na=False)
        L.append(f"\n## (2) '{kw}' matched: {int(mm.sum()):,} rows  Σ|amt|={a.abs()[mm].sum():,.0f}")
        if mm.any():
            g = (pd.DataFrame({"period": rp[mm], "proj": proj[mm], "theme": theme[mm], "v": vlab[mm], "amt": a.abs()[mm]})
                 .groupby(["period", "proj", "theme", "v"]).agg(n=("amt", "size"), amt=("amt", "sum"))
                 .sort_values("amt", ascending=False))
            for (p, pr, t, v), row in g.head(20).iterrows():
                L.append(f"   {str(p):9s} | {str(pr)[:26]:26s} | theme={str(t)[:8]:8s} | V={str(v)[:10]:10s} | {int(row['n']):>4} | {row['amt']/1e6:7.2f}M")

    # (3) per-period out-of-bucket
    try:
        sys.path.insert(0, str(ROOT / "src"))
        from kpi.lib.conf import load_categories
        from kpi.pipelines.step2_tag_projects._logic import normalize_ng_code
        cats = load_categories(); ng_cats = cats.get("ng_categories") or {}
        elig = {ng: (set(d.get("eligible_verticals") or []) | {"V_OTHER"}) for ng, d in ng_cats.items()}
        ngc = theme.map({s: (normalize_ng_code(s, cats) or "") for s in theme.unique()})
        oob = pd.Series(False, index=df.index)
        for ng in elig:
            oob |= ngc.eq(ng) & vid.ne("") & ~vid.isin(elig[ng])
        L.append("\n## (3) out-of-bucket (V not in NG eligible) per report_period — Σ|amt| M:")
        for p, x in a.abs()[oob].groupby(rp[oob]).sum().sort_values(ascending=False).items():
            L.append(f"   {str(p):9s} {x/1e6:9.2f}M  ({int((oob & rp.eq(p)).sum())} rows)")
        L.append(f"   TOTAL out-of-bucket: {int(oob.sum()):,} rows  {a.abs()[oob].sum()/1e6:.2f}M")
    except Exception as e:
        L.append(f"\n## (3) failed: {e}")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_sjm_v_check.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
