"""inspect_sjm_capex.py — sjm 25 capex -977 差額診斷
Run: python scripts\inspect_sjm_capex.py
Out: results\inspect_sjm_capex.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT   = Path(__file__).resolve().parent.parent
CSV    = ROOT / "tableau_combined_25.csv"
OUT    = ROOT / "results" / "inspect_sjm_capex.txt"
GOLDEN = 117736   # sjm 25 capex golden (萬)


def main():
    L = ["# inspect_sjm_capex — sjm 25 capex -977 差額診斷", ""]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return

    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    ent  = df["entity"].astype(str)
    yb   = df["year_bucket"].astype(str)
    y2   = yb.str[:2]
    co   = df.get("final_capex_opex", pd.Series("", index=df.index)).astype(str).str.strip()
    acct = df.get("account_code", pd.Series("", index=df.index)).astype(str).str.strip()
    adesc= df.get("account_desc", pd.Series("", index=df.index)).astype(str).str.strip()
    hid  = df.get("horizontal_id", pd.Series("", index=df.index)).astype(str).str.strip()

    def col(name):
        return pd.to_numeric(df.get(name, pd.Series(0, index=df.index)), errors="coerce").fillna(0.0)

    post = col("調整後_萬")
    pre  = col("調整前_萬")
    adj  = col("調整_萬")

    sjm25 = ent.eq("sjm") & y2.eq("25")
    sjm25c = sjm25 & co.str.lower().eq("capex")

    # ── 1. SJM 25 capex totals ──────────────────────────────────────────────────
    L += ["=" * 70, f"## 1. SJM 25 capex totals (golden={GOLDEN:,})"]
    L.append(f"  capex rows: {int(sjm25c.sum()):,}  調後Σ={post[sjm25c].sum():,.0f}  調前Σ={pre[sjm25c].sum():,.0f}  調Σ={adj[sjm25c].sum():,.0f}")
    L.append(f"  Δ (我哋 - golden): {post[sjm25c].sum()-GOLDEN:+,.0f}萬")

    # ── 2. By year_bucket ────────────────────────────────────────────────────────
    L += ["", "=" * 70, "## 2. SJM 25 capex by year_bucket"]
    L.append(f"  {'bucket':<12}  {'rows':>7}  {'調後_萬':>12}  {'調前_萬':>12}  {'調_萬':>12}")
    for bk in ["25", "25_24SY", "25_23SY"]:
        m = sjm25c & yb.eq(bk)
        L.append(f"  {bk:<12}  {int(m.sum()):>7,}  {post[m].sum():>12,.0f}  {pre[m].sum():>12,.0f}  {adj[m].sum():>12,.0f}")
    L.append(f"  {'TOTAL':<12}  {int(sjm25c.sum()):>7,}  {post[sjm25c].sum():>12,.0f}  {pre[sjm25c].sum():>12,.0f}  {adj[sjm25c].sum():>12,.0f}")

    # ── 3. SJM 25 ALL rows by capex/opex ────────────────────────────────────────
    L += ["", "=" * 70, "## 3. SJM 25 by final_capex_opex"]
    for v in ["Capex", "Opex", "", "nan"]:
        m = sjm25 & co.str.lower().eq(v.lower())
        if m.any():
            L.append(f"  [{v or 'BLANK'}]  {int(m.sum()):>7,} rows  調後Σ={post[m].sum():>12,.0f}  調前Σ={pre[m].sum():>12,.0f}")

    # ── 4. SJM 25 capex top accounts ────────────────────────────────────────────
    L += ["", "=" * 70, "## 4. SJM 25 capex by account (調後_萬, top 30)"]
    L.append(f"  {'account':<12} {'desc':<40} {'rows':>6}  {'調後_萬':>12}  {'調前_萬':>12}  {'調_萬':>10}")
    ag = df[sjm25c].copy()
    ag["_post"] = post[sjm25c].values
    ag["_pre"]  = pre[sjm25c].values
    ag["_adj"]  = adj[sjm25c].values
    ag["_acct"] = acct[sjm25c].values
    ag["_desc"] = adesc[sjm25c].values
    summary = (ag.groupby(["_acct","_desc"])
                 .agg(rows=("_post","count"), p=("_post","sum"),
                      q=("_pre","sum"), r=("_adj","sum"))
                 .sort_values("p", ascending=False).head(30))
    for (ac, ad), row in summary.iterrows():
        L.append(f"  {str(ac):<12} {str(ad)[:38]:<40} {int(row['rows']):>6,}  "
                 f"{row['p']:>12,.0f}  {row['q']:>12,.0f}  {row['r']:>10,.0f}")

    # ── 5. SJM 25 capex rows with NEGATIVE 調整後_萬 ────────────────────────────
    L += ["", "=" * 70, "## 5. SJM 25 capex rows with negative 調後_萬"]
    m_neg = sjm25c & (post < 0)
    L.append(f"  Negative rows: {int(m_neg.sum()):,}  Σ={post[m_neg].sum():,.0f}萬")
    if m_neg.any():
        ag_neg = df[m_neg].copy()
        ag_neg["_post"] = post[m_neg].values
        ag_neg["_pre"]  = pre[m_neg].values
        ag_neg["_adj"]  = adj[m_neg].values
        ag_neg["_acct"] = acct[m_neg].values
        ag_neg["_desc"] = adesc[m_neg].values
        s = (ag_neg.groupby(["_acct","_desc"])
                   .agg(rows=("_post","count"), p=("_post","sum"), r=("_adj","sum"))
                   .sort_values("p").head(15))
        for (ac, ad), r in s.iterrows():
            L.append(f"  {str(ac):<12} {str(ad)[:38]:<40} {int(r['rows']):>6,}  {r['p']:>12,.0f}  adj={r['r']:>10,.0f}")

    # ── 6. SJM 25 capex large 調整金額 rows ──────────────────────────────────────
    L += ["", "=" * 70, "## 6. SJM 25 capex: 調整金額 非零 rows"]
    m_adj = sjm25c & (adj.abs() > 0.001)
    L.append(f"  調整非零 rows: {int(m_adj.sum()):,}  Σ調={adj[m_adj].sum():,.0f}萬")
    if m_adj.any():
        ag_adj = df[m_adj].copy()
        ag_adj["_post"] = post[m_adj].values
        ag_adj["_adj"]  = adj[m_adj].values
        ag_adj["_acct"] = acct[m_adj].values
        ag_adj["_desc"] = adesc[m_adj].values
        s = (ag_adj.groupby(["_acct","_desc"])
                   .agg(rows=("_adj","count"), p=("_post","sum"), a=("_adj","sum"))
                   .sort_values("a").head(15))
        L.append(f"  {'account':<12} {'desc':<40} {'rows':>6}  {'調後Σ':>12}  {'調整Σ':>12}")
        for (ac, ad), r in s.iterrows():
            L.append(f"  {str(ac):<12} {str(ad)[:38]:<40} {int(r['rows']):>6,}  {r['p']:>12,.0f}  {r['a']:>12,.0f}")

    # ── 7. SJM 25 capex by H bucket ──────────────────────────────────────────────
    L += ["", "=" * 70, "## 7. SJM 25 capex by horizontal_id"]
    L.append(f"  {'H':<30} {'rows':>7}  {'調後_萬':>12}")
    for h, c in (df[sjm25c].assign(_post=post[sjm25c].values, _h=hid[sjm25c].values)
                            .groupby("_h")["_post"].agg(["count","sum"])
                            .sort_values("sum", ascending=False).iterrows()):
        L.append(f"  {str(h):<30} {int(c['count']):>7,}  {c['sum']:>12,.0f}")

    # ── 8. SJM 25 opex vs capex: can any opex be miscategorised capex? ──────────
    L += ["", "=" * 70, "## 8. SJM 25 opex rows Σ (cross-check)"]
    sjm25o = sjm25 & co.str.lower().eq("opex")
    L.append(f"  opex rows: {int(sjm25o.sum()):,}  調後Σ={post[sjm25o].sum():,.0f}  調前Σ={pre[sjm25o].sum():,.0f}")
    L.append(f"  capex rows: {int(sjm25c.sum()):,}  調後Σ={post[sjm25c].sum():,.0f}")
    L.append(f"  total 25: {int(sjm25.sum()):,}  調後Σ={post[sjm25].sum():,.0f}")

    # ── 9. Raw SJM capex from Excel (if accessible) ──────────────────────────────
    L += ["", "=" * 70, "## 9. SJM raw 'data' tab capex check"]
    F = ROOT / "data" / "sjm" / "raw" / "sjm_2025.xlsx"
    if F.exists():
        try:
            rdf = pd.read_excel(F, sheet_name="data", dtype=str)
            rdf.columns = [str(c).strip() for c in rdf.columns]
            L.append(f"  Loaded: {len(rdf):,} rows, cols={list(rdf.columns)[:20]}")
            # Find capex col
            cap_cols = [c for c in rdf.columns if "capex" in str(c).lower() or "報告" in str(c)]
            L.append(f"  Capex cols: {cap_cols}")
            for cc in cap_cols:
                vc = rdf[cc].astype(str).str.strip().value_counts().head(5)
                L.append(f"  [{cc}] top values: " + " | ".join(f"{v}:{c}" for v,c in vc.items()))
            # Sum 25跨年 for capex rows
            for cc in cap_cols:
                cap_mask = rdf[cc].astype(str).str.strip().str.lower().str.contains("capex", na=False)
                for amt_col in ["25跨年", "25調整金額"]:
                    if amt_col in rdf.columns:
                        num = pd.to_numeric(rdf.get(amt_col, 0), errors="coerce").fillna(0)
                        L.append(f"  [{cc}=Capex] {amt_col} Σ all={num.sum():,.0f}  Σ capex rows={num[cap_mask].sum():,.0f}")
        except Exception as e:
            L.append(f"  !! Error reading Excel: {e}")
    else:
        L.append(f"  SJM raw 揾唔到（{F}），跳過")

    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
