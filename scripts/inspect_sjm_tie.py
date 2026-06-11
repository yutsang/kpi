"""SJM — evaluate the NEW tie file (data/raw/sjm_2025.xlsx) and answer two questions:

(1) 新檔本身得唔得：row-level 調整前(amount_mop)/調整(adjustment_amount)/調整後(adjusted_amount)
    三個維度齊唔齊、逐行 tie 唔 tie（mop+adj=adjusted）、project/subproject map 咗未、
    25/25_24SY/25_23SY 每個 bucket 嘅 Σ(調整前/調整/調整後)。
(2) 舊資料有冇得救：同現有 pipeline 嘅 company_2_tagged_rows.parquet（25 buckets）對數 —
    bucket 總額對比 + project 級對比（top deltas）。如果 delta 集中喺幾個 project 嘅乾淨調整
    → 舊數據可以 patch 返；如果周身唔對 → 由新檔重建。
    （注意：舊 25 bucket 唔包 master-audit 階段先注入嘅 admin-comp 171 行/69.4M — 對數時心裡有數。）

Run (Windows):  python scripts/inspect_sjm_tie.py
Output: prints + results/inspect_sjm_tie.txt (small — paste back)
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parent.parent
NEW = ROOT / "data" / "raw" / "sjm_2025.xlsx"
TR = ROOT / "data" / "sjm" / "interim" / "company_2_tagged_rows.parquet"
CONF = ROOT / "conf" / "company_2" / "parameters.yml"
Y2B = {"2025": "25", "2024": "25_24SY", "2023": "25_23SY"}


def _num(s):
    return pd.to_numeric(s.astype("string").str.replace(",", "", regex=False), errors="coerce")


def _s(df, col):
    return df.get(col, pd.Series("", index=df.index)).astype("string").fillna("").str.strip()


def _normproj(s):
    return (s.astype("string").fillna("").str.strip().str.replace(r"\s+", "", regex=True)
            .str.replace("（", "(", regex=False).str.replace("）", ")", regex=False))


def main():
    L = ["# inspect_sjm_tie — new tie file vs old pipeline (row-level adjust dims + project tie)"]
    if not NEW.exists():
        L.append(f"X {NEW} missing"); _w(L); return
    xls = pd.ExcelFile(NEW)
    L.append(f"sheets: {xls.sheet_names} (reading first)")
    df = pd.read_excel(NEW, sheet_name=0, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    L.append(f"rows={len(df):,}  cols={len(df.columns)}")
    L.append(f"columns: {list(df.columns)}")

    # ── (1) row-level adjust dimensions ──────────────────────────────────────
    mop, adj, fin = _num(_s(df, "amount_mop")), _num(_s(df, "adjustment_amount")), _num(_s(df, "adjusted_amount"))
    has = {k: (k in df.columns) for k in ("amount_mop", "adjustment_amount", "adjusted_amount")}
    L.append(f"\n## (1) row-level 調整維度: {has}")
    L.append(f"   Σ調整前={mop.sum():,.0f}  Σ調整={adj.sum():,.0f}  Σ調整後={fin.sum():,.0f}")
    bad = ((mop.fillna(0) + adj.fillna(0)) - fin.fillna(0)).abs() > 1
    L.append(f"   逐行 tie (mop+adj=adjusted) 唔對嘅行: {int(bad.sum()):,}")
    blankf = fin.isna()
    L.append(f"   調整後空白行: {int(blankf.sum()):,}  (若全空=2025 慣例用調整前)")

    yr = _s(df, "year")
    L.append("\n## bucket 拆分 (year→bucket) — Σ調整前 / Σ調整 / Σ調整後 / rows:")
    for y, b in Y2B.items():
        m = yr.eq(y)
        if not int(m.sum()):
            L.append(f"   {b:8s} (year={y}): 0 rows"); continue
        L.append(f"   {b:8s} (year={y}): {int(m.sum()):>7,}r  前={mop[m].sum()/1e6:10.2f}M  "
                 f"調={adj[m].sum()/1e6:10.2f}M  後={fin[m].sum()/1e6:10.2f}M")
    oth = ~yr.isin(list(Y2B))
    if int(oth.sum()):
        L.append(f"   !! year 出 Y2B 範圍: {int(oth.sum())} rows  values={sorted(yr[oth].unique())[:8]}")

    # project / subproject mapping quality
    for c in ("project_code", "project", "subproject", "ng_theme", "capex_opex"):
        s = _s(df, c)
        L.append(f"   {c}: fill={s.ne('').mean()*100:.0f}%  distinct={s.nunique()}")
    pj = _s(df, "project")
    g = (pd.DataFrame({"p": pj.replace("", "(blank)"), "amt": fin.fillna(mop).fillna(0)})
         .groupby("p")["amt"].agg(["sum", "size"]))
    L.append("\n## top 12 projects (Σ調整後 or 前):")
    for p, r in g.sort_values("sum", key=lambda x: x.abs(), ascending=False).head(12).iterrows():
        L.append(f"   {str(p)[:44]:44s} {r['sum']/1e6:10.2f}M  ({int(r['size'])}r)")

    # ── (2) old pipeline comparison ───────────────────────────────────────────
    L.append("\n## (2) 舊 pipeline 對數 (company_2_tagged_rows.parquet)")
    if not TR.exists():
        L.append(f"   X {TR} missing — skip"); _w(L); return
    cfg = yaml.safe_load(CONF.read_text(encoding="utf-8")) if CONF.exists() else {}
    cols = (cfg or {}).get("columns") or {}
    old = pd.read_parquet(TR)
    amt_col = cols.get("amount") if cols.get("amount") in old.columns else (
        "amount" if "amount" in old.columns else None)
    pj_col = cols.get("project") if cols.get("project") in old.columns else None
    if not amt_col or "report_period" not in old.columns:
        L.append(f"   X cannot map amount/report_period (amount={amt_col!r})"); _w(L); return
    oa = pd.to_numeric(old[amt_col], errors="coerce").fillna(0.0)
    rp = old["report_period"].astype("string").fillna("")
    L.append(f"   old rows={len(old):,}  (admin-comp 69.4M 喺 audit 先注入,唔喺呢度)")
    L.append(f"   {'bucket':8s} {'OLD Σ':>12s} {'NEW Σ調整後':>12s} {'delta':>12s}")
    for y, b in Y2B.items():
        o = float(oa[rp.eq(b)].sum())
        nm = yr.eq(y)
        n = float(fin[nm].fillna(mop[nm]).sum())
        L.append(f"   {b:8s} {o/1e6:>11.2f}M {n/1e6:>11.2f}M {(n-o)/1e6:>11.2f}M")
    # project-level deltas (25 bucket only)
    if pj_col:
        op = _normproj(old.loc[rp.eq("25"), pj_col])
        oamt = oa[rp.eq("25")]
        og = pd.DataFrame({"p": op.replace("", "(blank)"), "amt": oamt}).groupby("p")["amt"].sum()
        nm = yr.eq("2025")
        ng = (pd.DataFrame({"p": _normproj(pj[nm]).replace("", "(blank)"),
                            "amt": fin[nm].fillna(mop[nm])}).groupby("p")["amt"].sum())
        j = pd.concat([og.rename("old"), ng.rename("new")], axis=1).fillna(0.0)
        j["delta"] = j["new"] - j["old"]
        j = j.reindex(j["delta"].abs().sort_values(ascending=False).index)
        L.append(f"\n## project 級 delta (bucket 25, |delta|>0.5M) — 集中=可patch舊, 周身=用新檔重建:")
        n_big = 0
        for p, r in j.iterrows():
            if abs(r["delta"]) < 0.5e6:
                continue
            n_big += 1
            if n_big <= 20:
                L.append(f"   {str(p)[:40]:40s} old={r['old']/1e6:9.2f}M new={r['new']/1e6:9.2f}M Δ={r['delta']/1e6:9.2f}M")
        L.append(f"   (|Δ|>0.5M projects: {n_big};  both-side matched projects: {int(((j['old']!=0)&(j['new']!=0)).sum())})")
    else:
        L.append(f"   (old project col {cols.get('project')!r} not in tagged_rows — project-level skip)")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_sjm_tie.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste back")


if __name__ == "__main__":
    main()
