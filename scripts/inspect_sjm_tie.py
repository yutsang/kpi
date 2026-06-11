"""SJM — evaluate the NEW tie file (data/raw/sjm_2025.xlsx, sheet 'data', NATIVE SAP cols).

The sjm tie file keeps original column names (NOT the unified 27-col schema):
  amount 調整前 = 'Val/COArea Crcy'   調整 = '調整金額'   調整後 = 前+調整 (computed row-level)
  year col = 'year' with values 25年項目 / 23年項目 / 跨年項目  (bucket mapping VERIFIED BY AMOUNT
  against golden: 25=124,661萬 25_24SY=63,594萬 25_23SY=7,971萬)
  project = 承批公司項目序號 + Project Name;  ng = 項目性質;  netoff = 'Net-off';  內部資源 col kept.

Reports:
 (1) row-level 前/調/後 totals + per-year-value Σ (so we SEE which value maps to which bucket)
 (2) netoff / 是否計入內部資源 slices (golden treatment differs)
 (3) old pipeline (tagged_rows) bucket totals vs golden + project-code level delta vs new file
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
GOLDEN = {"25": 124661e4, "25_24SY": 63594e4, "25_23SY": 7971e4}   # 萬 → MOP
# native-col candidates → canonical
CMAP = {
    "amount_mop": ["amount_mop", "Val/COArea Crcy"],
    "adjustment_amount": ["adjustment_amount", "調整金額"],
    "adjusted_amount": ["adjusted_amount"],
    "year": ["year", "項目期間", "年份"],
    "project_code": ["project_code", "承批公司項目序號"],
    "project": ["project", "Project Name"],
    "subproject": ["subproject", "Subproject"],
    "ng_theme": ["ng_theme", "項目性質"],
    "capex_opex": ["capex_opex", "報告capex/opex"],
    "netoff_flag": ["netoff_flag", "Net-off"],
    "internal": ["是否計入內部資源"],
    "unique_id": ["unique_id", "唯一識別碼"],
}


def _num(s):
    return pd.to_numeric(s.astype("string").str.replace(",", "", regex=False), errors="coerce")


def _s(df, col):
    return df.get(col, pd.Series("", index=df.index)).astype("string").fillna("").str.strip()


def main():
    L = ["# inspect_sjm_tie v2 — native SAP cols + golden-verified bucket mapping"]
    if not NEW.exists():
        L.append(f"X {NEW} missing"); _w(L); return
    xls = pd.ExcelFile(NEW)
    sh = next((s for s in xls.sheet_names if str(s).strip().lower() == "data"), xls.sheet_names[0])
    df = pd.read_excel(NEW, sheet_name=sh, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    ren = {}
    for canon, cands in CMAP.items():
        src = next((c for c in cands if c in df.columns), None)
        if src and src != canon:
            ren[src] = canon
    df = df.rename(columns=ren)
    L.append(f"sheet={sh!r}  rows={len(df):,}  renamed: {ren}")
    L.append(f"ALL columns: {list(df.columns)}")
    missing = [k for k in CMAP if k not in df.columns and k != "adjusted_amount"]
    L.append(f"canonical missing: {missing}")

    # ── duplication check (project totals ≈ 2× old → suspect stacked rows) ───
    uid = _s(df, "unique_id")
    L.append(f"\n## dup check: unique_id non-blank={int(uid.ne('').sum()):,} distinct={uid[uid.ne('')].nunique():,}")
    vc = uid[uid.ne("")].value_counts()
    dup_ids = vc[vc > 1]
    L.append(f"   unique_id 出現>1次: {len(dup_ids):,} ids / {int(dup_ids.sum()):,} rows")
    for i, (k, n) in enumerate(dup_ids.head(5).items()):
        L.append(f"      e.g. id={str(k)[:30]} ×{n}")
    full_dup = df.duplicated(keep=False)
    L.append(f"   exact full-row duplicates: {int(full_dup.sum()):,} rows")

    mop = _num(_s(df, "amount_mop")).fillna(0.0)
    adj = _num(_s(df, "adjustment_amount")).fillna(0.0)
    fin = (_num(_s(df, "adjusted_amount"))
           if "adjusted_amount" in df.columns else (mop + adj))
    fin = fin.fillna(mop + adj)
    L.append(f"\n## (1) row-level: Σ前={mop.sum()/1e6:,.2f}M  Σ調整={adj.sum()/1e6:,.2f}M  "
             f"Σ後={fin.sum()/1e6:,.2f}M   (後 = 前+調整, row-level ✓)")

    yr = _s(df, "year")
    net = _s(df, "netoff_flag")
    internal = _s(df, "internal")
    L.append("\n## per year-value Σ (前/調/後) — 用 golden 對返邊個 value 係邊個 bucket:")
    L.append(f"   golden:  25={GOLDEN['25']/1e6:,.2f}M  25_24SY={GOLDEN['25_24SY']/1e6:,.2f}M  "
             f"25_23SY={GOLDEN['25_23SY']/1e6:,.2f}M")
    for v in sorted(yr.unique()):
        m = yr.eq(v)
        L.append(f"   {str(v):12s} {int(m.sum()):>7,}r  前={mop[m].sum()/1e6:10.2f}M  "
                 f"調={adj[m].sum()/1e6:10.2f}M  後={fin[m].sum()/1e6:10.2f}M")
        # variants that golden might use: ex-netoff / ex-internal
        m2 = m & net.eq("") ; m3 = m2 & ~internal.str.upper().eq("Y")
        L.append(f"   {'':12s} {'ex-netoff':>12s}: 後={fin[m2].sum()/1e6:10.2f}M   "
                 f"ex-netoff+ex-內部資源: 後={fin[m3].sum()/1e6:10.2f}M")
    L.append(f"   netoff rows: {int(net.ne('').sum()):,} (Σ後={fin[net.ne('')].sum()/1e6:,.2f}M)   "
             f"內部資源=Y rows: {int(internal.str.upper().eq('Y').sum()):,} "
             f"(Σ後={fin[internal.str.upper().eq('Y')].sum()/1e6:,.2f}M)")

    for c in ("project_code", "project", "subproject", "ng_theme", "capex_opex"):
        s = _s(df, c)
        L.append(f"   {c}: fill={s.ne('').mean()*100:.0f}%  distinct={s.nunique()}")

    pc = _s(df, "project_code").str.replace(r"\.0$", "", regex=True)
    g = (pd.DataFrame({"p": pc.replace("", "(blank)") + "|" + _s(df, "project").str[:24], "amt": fin})
         .groupby("p")["amt"].agg(["sum", "size"]))
    L.append("\n## top 12 project_code|name (Σ後):")
    for p, r in g.sort_values("sum", key=lambda x: x.abs(), ascending=False).head(12).iterrows():
        L.append(f"   {str(p)[:48]:48s} {r['sum']/1e6:10.2f}M  ({int(r['size'])}r)")

    # ── (3) old pipeline ──────────────────────────────────────────────────────
    L.append("\n## (3) old pipeline tagged_rows vs golden / vs new (per project_code, bucket 25)")
    if not TR.exists():
        L.append(f"   X {TR} missing"); _w(L); return
    cfg = yaml.safe_load(CONF.read_text(encoding="utf-8")) if CONF.exists() else {}
    cols = (cfg or {}).get("columns") or {}
    old = pd.read_parquet(TR)
    amt_col = cols.get("amount") if cols.get("amount") in old.columns else "amount"
    if amt_col not in old.columns or "report_period" not in old.columns:
        L.append("   X old amount/report_period missing"); _w(L); return
    oa = pd.to_numeric(old[amt_col], errors="coerce").fillna(0.0)
    rp = old["report_period"].astype("string").fillna("")
    for b in ("25", "25_24SY", "25_23SY"):
        o = float(oa[rp.eq(b)].sum())
        L.append(f"   OLD {b:8s} {o/1e6:>10.2f}M   vs golden {GOLDEN[b]/1e6:>10.2f}M   Δ={(o-GOLDEN[b])/1e6:>8.2f}M")
    occ = next((c for c in old.columns if str(c).startswith("承批公司項目序號")), None)
    if occ:
        ocode = old[occ].astype("string").fillna("").str.strip().str.replace(r"\.0$", "", regex=True)
        og = pd.DataFrame({"p": ocode[rp.eq("25")].replace("", "(blank)"),
                           "amt": oa[rp.eq("25")]}).groupby("p")["amt"].sum()
        # new side: 25 bucket = whichever year-values the golden check says — show ALL for now
        ng = pd.DataFrame({"p": pc.replace("", "(blank)"), "amt": fin}).groupby("p")["amt"].sum()
        # dedup view: drop exact-duplicate rows, recompute per-code
        d2 = df.drop_duplicates()
        mop2 = _num(_s(d2, "amount_mop")).fillna(0.0)
        adj2 = _num(_s(d2, "adjustment_amount")).fillna(0.0)
        fin2 = mop2 + adj2
        pc2 = _s(d2, "project_code").str.replace(r"\.0$", "", regex=True)
        ng2 = pd.DataFrame({"p": pc2.replace("", "(blank)"), "amt": fin2}).groupby("p")["amt"].sum()
        L.append(f"\n   dedup: {len(df):,} → {len(d2):,} rows  Σ後 {fin.sum()/1e6:,.2f}M → {fin2.sum()/1e6:,.2f}M")
        j = pd.concat([og.rename("old25"), ng.rename("new_all"), ng2.rename("new_dedup")], axis=1).fillna(0.0)
        j["ratio"] = j["new_all"] / j["old25"].where(j["old25"] != 0)
        j["delta_dedup"] = j["new_dedup"] - j["old25"]
        j = j.reindex(j["new_all"].abs().sort_values(ascending=False).index)
        L.append("\n## project_code: old25 vs new vs new-dedup (ratio≈2 → stacked twice):")
        for p, r in j.head(15).iterrows():
            L.append(f"   code={str(p)[:8]:8s} old25={r['old25']/1e6:9.2f}M  new={r['new_all']/1e6:9.2f}M "
                     f"(×{r['ratio']:.3f})  dedup={r['new_dedup']/1e6:9.2f}M  Δdedup={r['delta_dedup']/1e6:8.2f}M")
        L.append(f"   matched codes(兩邊都有數): {int(((j['old25']!=0)&(j['new_all']!=0)).sum())} / old {int((j['old25']!=0).sum())} / new {int((j['new_all']!=0).sum())}")
    else:
        L.append("   (old 承批公司項目序號 col not found — project-level skip)")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_sjm_tie.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste back")


if __name__ == "__main__":
    main()
