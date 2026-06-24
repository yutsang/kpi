"""inspect_galaxy_comp_adj.py — galaxy 25 comp 調整前/後金額明細診斷
檢查 amount_mop / 調整金額 / 調整後金額 / 調整後_萬 / 調整前_萬 各欄
Run: python scripts\inspect_galaxy_comp_adj.py
Out: results\inspect_galaxy_comp_adj.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV  = ROOT / "tableau_combined_25.csv"
OUT  = ROOT / "results" / "inspect_galaxy_comp_adj.txt"
COMP_H = {"H_HOTEL_ROOM", "H_VENUE", "H_FNB", "H_COMP_TICKET", "H_COMP_OTHER"}


def main():
    L = ["# inspect_galaxy_comp_adj — galaxy 25 comp 調整前/後金額診斷", ""]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return

    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)

    # Amount columns present
    amt_cols = ["amount_mop", "調整金額", "調整後金額", "調整後_萬", "調整前_萬", "調整_萬"]
    L += ["=" * 70, "## 0. Amount columns present in CSV"]
    for c in amt_cols:
        if c in df.columns:
            ser = pd.to_numeric(df[c], errors="coerce")
            nb = ser.notna().sum()
            nz = (ser != 0).sum()
            L.append(f"  {c:<20} present  non-null={nb:,}  non-zero={nz:,}")
        else:
            L.append(f"  {c:<20} !! NOT IN CSV")

    # Filter galaxy 25 comp
    ent = df["entity"].astype(str)
    yb  = df["year_bucket"].astype(str)
    y2  = yb.str[:2]
    hid = df.get("horizontal_id", pd.Series("", index=df.index)).astype(str).str.strip()
    acct = df.get("account_code", pd.Series("", index=df.index)).astype(str).str.strip()
    adesc = df.get("account_desc", pd.Series("", index=df.index)).astype(str).str.strip()

    def col(name):
        return pd.to_numeric(df.get(name, pd.Series(0, index=df.index)), errors="coerce").fillna(0.0)

    post = col("調整後_萬")
    pre  = col("調整前_萬")
    adj  = col("調整_萬")
    amop = col("amount_mop")
    aadj = col("調整後金額")
    kadj = col("調整金額")

    gx25  = ent.eq("galaxy") & y2.eq("25")
    comp25 = gx25 & hid.isin(COMP_H)

    # ── 1. Galaxy 25 comp totals by column ──────────────────────────────────────
    L += ["", "=" * 70, "## 1. galaxy 25 comp — totals by amount column (golden=59,828)"]
    for name, s in [("amount_mop/1e4", amop[comp25]/1e4),
                    ("調整後_萬 (我哋用)", post[comp25]),
                    ("調整前_萬",          pre[comp25]),
                    ("調整_萬",            adj[comp25]),
                    ("調整後金額/1e4",      aadj[comp25]/1e4),
                    ("調整金額/1e4",        kadj[comp25]/1e4)]:
        L.append(f"  {name:<25}  Σ={s.sum():>12,.2f}萬  non-zero rows={int((s!=0).sum()):,}")
    L.append(f"  {'golden':25}  Σ={59828:>12,}萬")

    # ── 2. 調整前 vs 調整後 by H bucket ──────────────────────────────────────────
    L += ["", "=" * 70, "## 2. galaxy 25 comp: 調整前 vs 調整後 by H bucket"]
    L.append(f"  {'H bucket':<22}  {'rows':>7}  {'調整前_萬':>12}  {'調整後_萬':>12}  {'調整_萬':>12}")
    for h in sorted(COMP_H):
        m = comp25 & hid.eq(h)
        if not m.any(): continue
        L.append(f"  {h:<22}  {m.sum():>7,}  {pre[m].sum():>12,.0f}  {post[m].sum():>12,.0f}  {adj[m].sum():>12,.0f}")
    L.append(f"  {'TOTAL':<22}  {comp25.sum():>7,}  {pre[comp25].sum():>12,.0f}  {post[comp25].sum():>12,.0f}  {adj[comp25].sum():>12,.0f}")

    # ── 3. Top accounts: 調整前 vs 調整後 ────────────────────────────────────────
    L += ["", "=" * 70, "## 3. top 20 accounts in galaxy 25 comp: 調整前 vs 調整後"]
    L.append(f"  {'account':<15} {'desc'[:35]:<37} {'rows':>6}  {'調整前':>11}  {'調整後':>11}  {'調整額':>11}")
    ag = df[comp25].copy()
    ag["_pre"]  = pre[comp25].values
    ag["_post"] = post[comp25].values
    ag["_adj"]  = adj[comp25].values
    ag["_acct"] = acct[comp25].values
    ag["_desc"] = adesc[comp25].values
    summary = (ag.groupby(["_acct", "_desc"])
                 .agg(rows=("_pre","count"), pre_sum=("_pre","sum"),
                      post_sum=("_post","sum"), adj_sum=("_adj","sum"))
                 .sort_values("post_sum", ascending=False).head(20))
    for (ac, ad), row in summary.iterrows():
        L.append(f"  {str(ac):<15} {str(ad)[:35]:<37} {int(row['rows']):>6,}  "
                 f"{row['pre_sum']:>11,.0f}  {row['post_sum']:>11,.0f}  {row['adj_sum']:>11,.0f}")

    # ── 4. 全家 comp 調整前 vs 調整後 vs golden ──────────────────────────────────
    L += ["", "=" * 70, "## 4. 全家 25年 comp: 調整前 vs 調整後 vs golden"]
    goldens = {"galaxy":59828,"wynn":19823,"vml":8396,"melco":8403,"mgm":7532,"sjm":2920}
    L.append(f"  {'entity':<8} {'golden':>10} {'調整前':>12} {'Δpre':>9} {'調整後':>12} {'Δpost':>9}")
    for e, g in goldens.items():
        m = ent.eq(e) & y2.eq("25") & hid.isin(COMP_H)
        L.append(f"  {e:<8} {g:>10,} {pre[m].sum():>12,.0f} {pre[m].sum()-g:>+9,.0f} "
                 f"{post[m].sum():>12,.0f} {post[m].sum()-g:>+9,.0f}")

    # ── 5. 8107350 Venue License — 調整前/後 row-level sample ────────────────────
    L += ["", "=" * 70, "## 5. 8107350 Venue License — row sample (galaxy 25, first 20 rows)"]
    m8 = comp25 & acct.eq("8107350")
    n8 = int(m8.sum())
    L.append(f"  total rows={n8:,}  調整前 Σ={pre[m8].sum():,.0f}  調整後 Σ={post[m8].sum():,.0f}  調整 Σ={adj[m8].sum():,.0f}")
    sub8 = df[m8].copy()
    sub8["_pre"]  = pre[m8].values
    sub8["_post"] = post[m8].values
    sub8["_adj"]  = adj[m8].values
    sub8["_amop"] = amop[m8].values
    sub8["_aadj"] = aadj[m8].values
    sub8["_kadj"] = kadj[m8].values
    L.append(f"  {'調前_萬':>10} {'調後_萬':>10} {'調_萬':>10} {'amop/1e4':>10} {'調後金額/1e4':>12} {'調金額/1e4':>12}")
    for _, r in sub8.head(20).iterrows():
        L.append(f"  {r['_pre']:>10,.2f} {r['_post']:>10,.2f} {r['_adj']:>10,.2f} "
                 f"{r['_amop']/1e4:>10,.2f} {r['_aadj']/1e4:>12,.2f} {r['_kadj']/1e4:>12,.2f}")

    # ── 6. 7120100 Comp Lodging — 調整前/後 row-level sample ────────────────────
    L += ["", "=" * 70, "## 6. 7120100 Comp Lodging — row sample (galaxy 25, first 20 rows)"]
    m71 = comp25 & acct.eq("7120100")
    n71 = int(m71.sum())
    L.append(f"  total rows={n71:,}  調整前 Σ={pre[m71].sum():,.0f}  調整後 Σ={post[m71].sum():,.0f}  調整 Σ={adj[m71].sum():,.0f}")
    sub71 = df[m71].copy()
    sub71["_pre"]  = pre[m71].values; sub71["_post"] = post[m71].values
    sub71["_adj"]  = adj[m71].values; sub71["_amop"] = amop[m71].values
    L.append(f"  {'調前_萬':>10} {'調後_萬':>10} {'調_萬':>10} {'amop/1e4':>10}")
    for _, r in sub71.head(20).iterrows():
        L.append(f"  {r['_pre']:>10,.4f} {r['_post']:>10,.4f} {r['_adj']:>10,.4f} {r['_amop']/1e4:>10,.4f}")

    # ── 7. galaxy 25 non-comp rows: 調整後金額 populated? ────────────────────────
    L += ["", "=" * 70, "## 7. galaxy 25 rows: 調整後金額 vs amount_mop comparison"]
    gx25all = ent.eq("galaxy") & y2.eq("25")
    aadj_gx25 = aadj[gx25all]
    amop_gx25 = amop[gx25all]
    has_aadj = (aadj_gx25 != 0).sum()
    same = ((aadj_gx25 - amop_gx25).abs() < 1).sum()
    L.append(f"  galaxy 25 rows: {int(gx25all.sum()):,}")
    L.append(f"  調後金額 non-zero: {int(has_aadj):,}  (of {int(gx25all.sum()):,})")
    L.append(f"  調後金額 ≈ amount_mop (diff<1): {int(same):,}")
    L.append(f"  調後_萬 Σ (all gx25) = {post[gx25all].sum():,.0f}")
    L.append(f"  調前_萬 Σ (all gx25) = {pre[gx25all].sum():,.0f}")
    L.append(f"  調_萬   Σ (all gx25) = {adj[gx25all].sum():,.0f}")

    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
