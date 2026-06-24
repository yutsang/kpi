"""inspect_sjm_capex3.py — sjm raw 調整事項(Y) rows + "zeroed out" rows 分析
找出所有「bucket金額被自己調整金額完全對沖→淨=0」的capex行，以及 golden 可能如何計佢哋
Run: python scripts\inspect_sjm_capex3.py
Out: results\inspect_sjm_capex3.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
F    = ROOT / "data" / "sjm" / "raw" / "sjm_2025.xlsx"
OUT  = ROOT / "results" / "inspect_sjm_capex3.txt"
GOLDEN = 117736   # sjm 25 capex golden (萬)


def _num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s.astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0.0)


def main():
    L = ["# inspect_sjm_capex3 — sjm raw 調整事項(Y) + zeroed-out rows", ""]
    if not F.exists():
        L.append(f"!! {F} 揾唔到"); _w(L); return

    rdf = pd.read_excel(F, sheet_name="data", dtype=str)
    rdf.columns = [str(c).strip() for c in rdf.columns]
    n = len(rdf)
    L.append(f"Loaded: {n:,} rows, {len(rdf.columns)} cols")

    cap_mask = rdf.get("報告capex/opex", pd.Series("", index=rdf.index)).astype(str).str.strip().str.lower().eq("capex")
    L.append(f"Capex rows: {int(cap_mask.sum()):,}")

    # numeric cols
    amt_cols = {c: _num(rdf[c]) for c in ["25跨年","25調整金額","24跨年","24調整金額","23跨年","23調整金額","Val/COArea Crcy"]
                if c in rdf.columns}
    net25 = amt_cols.get("25跨年", pd.Series(0.0, index=rdf.index)) + amt_cols.get("25調整金額", pd.Series(0.0, index=rdf.index))
    net24 = amt_cols.get("24跨年", pd.Series(0.0, index=rdf.index)) + amt_cols.get("24調整金額", pd.Series(0.0, index=rdf.index))
    net23 = amt_cols.get("23跨年", pd.Series(0.0, index=rdf.index)) + amt_cols.get("23調整金額", pd.Series(0.0, index=rdf.index))
    net_total = net25 + net24 + net23
    val = amt_cols.get("Val/COArea Crcy", pd.Series(0.0, index=rdf.index))

    # ── 1. 調整事項(Y) distribution ────────────────────────────────────────────
    L += ["", "=" * 70, "## 1. 調整事項(Y) value counts (all rows / capex rows)"]
    yf = "調整事項(Y)"
    if yf in rdf.columns:
        vc_all = rdf[yf].astype(str).str.strip().value_counts(dropna=False)
        vc_cap = rdf[cap_mask][yf].astype(str).str.strip().value_counts(dropna=False)
        L.append(f"  {'value':<15}  {'all rows':>10}  {'capex rows':>10}")
        for v in vc_all.index[:10]:
            L.append(f"  {str(v):<15}  {int(vc_all.get(v,0)):>10,}  {int(vc_cap.get(v,0)):>10,}")
        y_mask = rdf[yf].astype(str).str.strip().str.upper().eq("Y")
        L.append(f"\n  Y rows: {int(y_mask.sum()):,} all / {int((y_mask&cap_mask).sum()):,} capex")
        L.append(f"  Y rows (capex) Val/COArea Σ = {val[y_mask & cap_mask].sum()/1e4:,.2f}萬")
        L.append(f"  Y rows (capex) net25+24+23  Σ = {net_total[y_mask & cap_mask].sum()/1e4:,.2f}萬")
    else:
        L.append(f"  調整事項(Y) col not found")

    # ── 2. 全部「zeroed-out」capex rows — net = 0 but Val/COArea ≠ 0 ─────────
    L += ["", "=" * 70, "## 2. capex rows where net_all_buckets = 0 but Val/COArea ≠ 0"]
    zeroed = cap_mask & (abs(net_total) < 1) & (abs(val) > 1)
    L.append(f"  Rows: {int(zeroed.sum()):,}  Val/COArea Σ = {val[zeroed].sum()/1e4:,.2f}萬")
    if zeroed.any():
        sub = rdf[zeroed].copy()
        sub["_val"] = val[zeroed].values
        sub["_net25"] = net25[zeroed].values
        sub["_net24"] = net24[zeroed].values
        sub["_net23"] = net23[zeroed].values
        # show by 調整事項(Y) flag
        if "調整事項(Y)" in sub.columns:
            for yv in ["Y","nan","N",""]:
                m = sub["調整事項(Y)"].astype(str).str.strip().str.upper().eq(yv.upper())
                c = int(m.sum())
                if c > 0:
                    L.append(f"  調整事項(Y)={yv!r}: {c} rows  Val Σ={sub[m]['_val'].sum()/1e4:,.2f}萬")
        # top accounts
        L += ["", "  Top accounts in zeroed-out capex:"]
        L.append(f"  {'account':<12}  {'desc':<40}  {'rows':>6}  {'Val_萬':>10}")
        ce = "Cost Element" if "Cost Element" in rdf.columns else None
        cd = "Cost element descr." if "Cost element descr." in rdf.columns else None
        if ce and cd:
            sub["_ce"] = rdf[zeroed][ce].astype(str).values
            sub["_cd"] = rdf[zeroed][cd].astype(str).values
            grp = (sub.groupby(["_ce","_cd"])["_val"].agg(["count","sum"])
                      .sort_values("sum", ascending=False).head(20))
            for (c_, d_), r in grp.iterrows():
                L.append(f"  {str(c_):<12}  {str(d_)[:38]:<40}  {int(r['count']):>6,}  {r['sum']/1e4:>10,.2f}")

    # ── 3. Total golden recon attempt: add zeroed-out Val/COArea ────────────────
    L += ["", "=" * 70, "## 3. Recon attempt: standard 3-bucket + zeroed-out Val/COArea"]
    standard = net_total[cap_mask].sum() / 1e4
    zeroed_val = val[zeroed].sum() / 1e4
    attempt = standard + zeroed_val
    L.append(f"  Standard 3-bucket (capex): {standard:,.2f}萬")
    L.append(f"  + zeroed-out Val/COArea:   {zeroed_val:,.2f}萬")
    L.append(f"  = Attempt total:           {attempt:,.2f}萬  (golden={GOLDEN:,}  Δ={(attempt-GOLDEN):+,.2f})")

    # ── 4. Rows with partial zeroing: net ≠ 0 but |val| >> |net| ────────────────
    L += ["", "=" * 70, "## 4. capex rows with large adjustment (|adj/amt| > 50%) — sample"]
    adj25 = amt_cols.get("25調整金額", pd.Series(0.0, index=rdf.index))
    amt25 = amt_cols.get("25跨年", pd.Series(0.0, index=rdf.index))
    large_adj = cap_mask & (abs(adj25) > abs(amt25) * 0.5) & (abs(amt25) > 0)
    L.append(f"  Rows with |25調整|>50%×25跨年: {int(large_adj.sum()):,}")
    L.append(f"  Σ 25跨年={amt25[large_adj].sum()/1e4:,.2f}  Σ 25調整={adj25[large_adj].sum()/1e4:,.2f}  Σ net25={net25[large_adj].sum()/1e4:,.2f}萬")

    # ── 5. 23年期後 column ────────────────────────────────────────────────────────
    L += ["", "=" * 70, "## 5. 23年期後 column (capex rows)"]
    pc = "23年期後"
    if pc in rdf.columns:
        vc = rdf[cap_mask][pc].astype(str).str.strip().value_counts(dropna=False).head(10)
        L.append(f"  value counts (capex):")
        for v, c_ in vc.items():
            L.append(f"    {str(v):<20}: {c_:,}")
        y_mask23 = rdf[pc].astype(str).str.strip().str.upper().eq("Y")
        L.append(f"  Y rows (capex): {int((y_mask23 & cap_mask).sum()):,}")
        if (y_mask23 & cap_mask).any():
            L.append(f"  Val/COArea Σ = {val[y_mask23 & cap_mask].sum()/1e4:,.2f}萬")
            L.append(f"  net_total  Σ = {net_total[y_mask23 & cap_mask].sum()/1e4:,.2f}萬")
    else:
        L.append(f"  23年期後 col not found")

    # ── 6. 哪些 capex rows 在 golden 多出嚟？Brute-force: val ─ net gaps per row ──
    L += ["", "=" * 70, "## 6. Val/COArea vs net_total diff for capex rows"]
    diff = (val - net_total)
    diff_cap = diff[cap_mask]
    L.append(f"  Σ(Val-net) for capex rows = {diff_cap.sum()/1e4:,.2f}萬")
    L.append(f"  Σ(Val-net) > 0 rows: {int((diff_cap > 1).sum()):,}  Σ={diff_cap[diff_cap>1].sum()/1e4:,.2f}萬")
    L.append(f"  Σ(Val-net) < 0 rows: {int((diff_cap < -1).sum()):,}  Σ={diff_cap[diff_cap<-1].sum()/1e4:,.2f}萬")

    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
