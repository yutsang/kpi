r"""MGM comp H 行按 source 欄位分組，對比 PM/OPEX-PM/WD5-Patron 金額。
User raw Excel: 23=6355 / 24=3140 / 25=3210（含各前綴，調整後/調整前）

Run（Windows）:
  python scripts\inspect_mgm_pm_source.py
出 results\inspect_mgm_pm_source.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
COMP_H = {"H_HOTEL_ROOM", "H_VENUE", "H_FNB", "H_COMP_TICKET", "H_COMP_OTHER"}

# User raw Excel reference
USER_PM = {"23": 6355, "24": 3140, "25": 3210}


def _amt_row(row_yb, post, pre):
    y2 = row_yb.str[:2]
    return post.where(y2.isin(["23", "24"]), pre)


def main():
    L = ["# inspect_mgm_pm_source — MGM comp 按 source 分組（找 PM vs 非PM，對比 raw）"]
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    mg = df[df["entity"].astype(str) == "mgm"].copy()

    post = pd.to_numeric(mg.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    pre  = pd.to_numeric(mg.get("調整前_萬",  0), errors="coerce").fillna(0.0)
    yb   = mg["year_bucket"].astype(str)
    y2   = yb.str[:2]
    mg["m"] = _amt_row(yb, post, pre)
    mg["bk_pref"] = y2

    hid = mg["horizontal_id"].astype(str).str.strip() if "horizontal_id" in mg.columns else pd.Series("", index=mg.index)
    src = mg["source"].astype(str).str.strip()        if "source"        in mg.columns else pd.Series("", index=mg.index)
    ad  = mg["account_desc"].astype(str)              if "account_desc"  in mg.columns else pd.Series("", index=mg.index)

    comp_mask = hid.isin(COMP_H)

    # ── A: source 唯一值 in comp H rows (all buckets) ──
    L.append("\n## A: source 所有唯一值（comp H 行）+ Σ")
    g_src = mg[comp_mask].groupby(src.reindex(mg[comp_mask].index))["m"].sum().reset_index()
    g_src.columns = ["source", "Σ"]
    g_src = g_src.sort_values("Σ", key=lambda s: s.abs(), ascending=False)
    for _, r in g_src.head(30).iterrows():
        L.append(f"  {str(r['source']):<30} {r['Σ']:>10,.0f}萬")

    # ── B: 逐年（prefix）PM vs 非PM ──
    PM_KW = ["PM", "OPEX-PM", "WD5-Patron", "WD5", "Patron"]
    def is_pm(s): return any(k.lower() in s.lower() for k in PM_KW)
    pm_mask = src.apply(is_pm)

    L.append(f"\n{'='*64}")
    L.append("## B: 逐年（spent-year prefix） PM vs 非PM comp Σ")
    L.append(f"  {'bucket':<8} {'PM Σ':>10} {'user ref':>10} {'Δ':>8} {'非PM Σ':>10} {'comp合計':>10}")
    for yr in ("23", "24", "25"):
        if yr == "23":
            ym = yb.eq("23")
        else:
            ym = y2.eq(yr)
        sub_comp = mg[comp_mask & ym]
        sub_src  = src.reindex(sub_comp.index)
        pm_sub   = sub_comp[pm_mask.reindex(sub_comp.index, fill_value=False)]
        npm_sub  = sub_comp[~pm_mask.reindex(sub_comp.index, fill_value=False)]
        pm_sum   = pm_sub["m"].sum()
        npm_sum  = npm_sub["m"].sum()
        total    = sub_comp["m"].sum()
        ref      = USER_PM.get(yr, 0)
        L.append(f"  {yr:<8} {pm_sum:>10,.0f} {ref:>10,.0f} {pm_sum-ref:>8,.0f} {npm_sum:>10,.0f} {total:>10,.0f}")

    # ── C: PM sub-source breakdown by year ──
    L.append(f"\n{'='*64}\n## C: PM source 細項 per year（bucket prefix）")
    for yr in ("23", "24", "25"):
        if yr == "23":
            ym = yb.eq("23")
        else:
            ym = y2.eq(yr)
        sub_pm = mg[comp_mask & ym & pm_mask]
        if not len(sub_pm):
            L.append(f"\n  bucket {yr}: 無 PM comp 行")
            continue
        L.append(f"\n  bucket {yr} PM comp Σ={sub_pm['m'].sum():,.0f}萬 rows={len(sub_pm):,}")
        g2 = sub_pm.groupby([src.reindex(sub_pm.index), "account_desc"])["m"].sum().reset_index()
        g2 = g2[g2["m"].abs() >= 10].sort_values("m", key=lambda s: s.abs(), ascending=False)
        for _, r in g2.head(15).iterrows():
            L.append(f"    src={str(r['source']):<18} {str(r['account_desc'])[:30]:<30} {r['m']:>8,.0f}萬")

    # ── D: 非PM comp breakdown（找應搬走嘅行）──
    L.append(f"\n{'='*64}\n## D: 非PM comp 行 breakdown（考慮搬走）")
    for yr in ("23", "24", "25"):
        if yr == "23":
            ym = yb.eq("23")
        else:
            ym = y2.eq(yr)
        sub_npm = mg[comp_mask & ym & ~pm_mask]
        if not len(sub_npm):
            L.append(f"\n  bucket {yr}: 無非PM comp 行")
            continue
        L.append(f"\n  bucket {yr} 非PM comp Σ={sub_npm['m'].sum():,.0f}萬 rows={len(sub_npm):,}")
        g3 = sub_npm.groupby([src.reindex(sub_npm.index), "account_desc"])["m"].sum().reset_index()
        g3 = g3[g3["m"].abs() >= 50].sort_values("m", key=lambda s: s.abs(), ascending=False)
        for _, r in g3.head(15).iterrows():
            L.append(f"    src={str(r['source']):<18} {str(r['account_desc'])[:30]:<30} {r['m']:>8,.0f}萬")

    # ── E: sjm comp 分析（用 account_desc 搵可清嘅行）──
    L.append(f"\n{'='*64}\n## E: sjm comp 逐年（搵可清行）")
    sj = df[df["entity"].astype(str) == "sjm"].copy()
    post_s = pd.to_numeric(sj.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    pre_s  = pd.to_numeric(sj.get("調整前_萬",  0), errors="coerce").fillna(0.0)
    yb_s = sj["year_bucket"].astype(str); y2_s = yb_s.str[:2]
    sj["m"] = _amt_row(yb_s, post_s, pre_s)
    hid_s = sj["horizontal_id"].astype(str).str.strip() if "horizontal_id" in sj.columns else pd.Series("", index=sj.index)
    comp_s = hid_s.isin(COMP_H)
    HQ_SJM = {"23": 2970, "24": 6553, "25": 3575}
    for yr in ("23", "24", "25"):
        if yr == "23":
            ym = yb_s.eq("23")
        else:
            ym = y2_s.eq(yr)
        sub = sj[comp_s & ym]
        total = sub["m"].sum()
        hq = HQ_SJM.get(yr, 0)
        L.append(f"\n  sjm bucket {yr}  Σ={total:,.0f}萬  HQ={hq:,.0f}萬  Δ={total-hq:,.0f}萬")
        if not len(sub): continue
        ad_s = sub["account_desc"].astype(str) if "account_desc" in sub.columns else pd.Series("", index=sub.index)
        g = sub.groupby(["horizontal_id", ad_s])["m"].sum().reset_index()
        g = g[g["m"].abs() >= 30].sort_values("m", key=lambda s: s.abs(), ascending=False)
        for _, r in g.head(12).iterrows():
            L.append(f"    {r['horizontal_id']:<14} {str(r['account_desc'])[:34]:<34} {r['m']:>8,.0f}萬")

    out = ROOT / "results" / "inspect_mgm_pm_source.txt"
    out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
