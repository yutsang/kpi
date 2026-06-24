"""inspect_galaxy_staff.py — galaxy 25 staff +6,669 帳號明細診斷
Run: python scripts\inspect_galaxy_staff.py
Out: results\inspect_galaxy_staff.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT   = Path(__file__).resolve().parent.parent
CSV    = ROOT / "tableau_combined_25.csv"
OUT    = ROOT / "results" / "inspect_galaxy_staff.txt"
GOLDEN = 14568   # galaxy 25 staff golden (萬)


def main():
    L = ["# inspect_galaxy_staff — galaxy 25 staff +6,669 帳號診斷", ""]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return

    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    ent  = df["entity"].astype(str)
    yb   = df["year_bucket"].astype(str)
    y2   = yb.str[:2]
    hid  = df.get("horizontal_id", pd.Series("", index=df.index)).astype(str).str.strip()
    acct = df.get("account_code", pd.Series("", index=df.index)).astype(str).str.strip()
    adesc = df.get("account_desc", pd.Series("", index=df.index)).astype(str).str.strip()
    desc  = df.get("description", pd.Series("", index=df.index)).astype(str).str.strip()
    co    = df.get("final_capex_opex", pd.Series("", index=df.index)).astype(str).str.strip()

    def col(name):
        return pd.to_numeric(df.get(name, pd.Series(0, index=df.index)), errors="coerce").fillna(0.0)

    post = col("調整後_萬")
    pre  = col("調整前_萬")
    adj  = col("調整_萬")
    amop = col("amount_mop")

    gx25_staff = ent.eq("galaxy") & y2.eq("25") & hid.eq("H_LABOR")
    gx24_staff = ent.eq("galaxy") & y2.eq("24") & hid.eq("H_LABOR")

    # ── 1. Totals ────────────────────────────────────────────────────────────────
    L += ["=" * 70, f"## 1. galaxy staff totals (golden 25={GOLDEN:,})"]
    for yr, mask in [("25", gx25_staff), ("24", gx24_staff)]:
        L.append(f"  {yr}: {int(mask.sum()):>8,} rows  調後Σ={post[mask].sum():>10,.0f}  "
                 f"調前Σ={pre[mask].sum():>10,.0f}  調Σ={adj[mask].sum():>10,.0f}")
    L.append(f"  Δ (25 - golden): {post[gx25_staff].sum()-GOLDEN:+,.0f}萬")

    # ── 2. By account code: 調後 25 vs 24 ────────────────────────────────────────
    L += ["", "=" * 70, "## 2. galaxy staff 25 by account (調後_萬, top 30)"]
    L.append(f"  {'account':<12} {'desc':<40} {'rows':>6}  {'調後25':>10}  {'調後24':>10}  {'excess':>10}  {'調25':>10}")
    ag25 = df[gx25_staff].copy()
    ag25["_post"] = post[gx25_staff].values
    ag25["_adj"]  = adj[gx25_staff].values
    ag25["_acct"] = acct[gx25_staff].values
    ag25["_desc"] = adesc[gx25_staff].values
    sum25 = (ag25.groupby(["_acct","_desc"])
               .agg(rows=("_post","count"), post25=("_post","sum"), adj25=("_adj","sum"))
               .reset_index())

    ag24 = df[gx24_staff].copy()
    ag24["_post"] = post[gx24_staff].values
    ag24["_acct"] = acct[gx24_staff].values
    ag24["_desc"] = adesc[gx24_staff].values
    sum24 = (ag24.groupby(["_acct","_desc"])
               .agg(post24=("_post","sum"))
               .reset_index())

    merged = sum25.merge(sum24, on=["_acct","_desc"], how="left").fillna({"post24": 0.0})
    merged["excess"] = merged["post25"] - merged["post24"]
    merged = merged.sort_values("post25", ascending=False).head(30)
    for _, r in merged.iterrows():
        L.append(f"  {str(r['_acct']):<12} {str(r['_desc'])[:38]:<40} "
                 f"{int(r['rows']):>6,}  {r['post25']:>10,.0f}  {r['post24']:>10,.0f}  "
                 f"{r['excess']:>10,.0f}  {r['adj25']:>10,.0f}")

    # ── 3. Capex vs Opex breakdown ────────────────────────────────────────────────
    L += ["", "=" * 70, "## 3. galaxy 25 staff: Capex vs Opex"]
    for cap in ["Capex", "Opex"]:
        m = gx25_staff & co.str.lower().eq(cap.lower())
        L.append(f"  {cap}: {int(m.sum()):>7,} rows  調後Σ={post[m].sum():>10,.0f}  調Σ={adj[m].sum():>10,.0f}")

    # ── 4. Accounts with large 調整 (negative) → should reduce staff ─────────────
    L += ["", "=" * 70, "## 4. galaxy 25 staff: accounts with non-zero 調整金額"]
    m_adj = gx25_staff & (adj.abs() > 0.001)
    L.append(f"  調整金額非0 rows: {int(m_adj.sum()):,}  Σ調整={adj[m_adj].sum():,.0f}萬")
    if m_adj.any():
        ag_adj = df[m_adj].copy()
        ag_adj["_adj"]  = adj[m_adj].values
        ag_adj["_post"] = post[m_adj].values
        ag_adj["_acct"] = acct[m_adj].values
        ag_adj["_desc"] = adesc[m_adj].values
        s = (ag_adj.groupby(["_acct","_desc"])
                   .agg(rows=("_adj","count"), post_sum=("_post","sum"), adj_sum=("_adj","sum"))
                   .sort_values("adj_sum").head(20))
        L.append(f"  {'account':<12} {'desc':<40} {'rows':>6}  {'調後Σ':>10}  {'調整Σ':>10}")
        for (ac, ad), r in s.iterrows():
            L.append(f"  {str(ac):<12} {str(ad)[:38]:<40} {int(r['rows']):>6,}  "
                     f"{r['post_sum']:>10,.0f}  {r['adj_sum']:>10,.0f}")

    # ── 5. 8000960 focus ──────────────────────────────────────────────────────────
    L += ["", "=" * 70, "## 5. Account 8000960 detail (Payroll-Direct Event Investment)"]
    m89 = ent.eq("galaxy") & acct.eq("8000960")
    for yr, mask_yr in [("25", y2.eq("25")), ("24", y2.eq("24"))]:
        m = m89 & mask_yr
        m_staff = m & hid.eq("H_LABOR")
        m_other = m & ~hid.eq("H_LABOR")
        L.append(f"  {yr}: total={int(m.sum()):,} rows  staff={int(m_staff.sum()):,}/調後{post[m_staff].sum():,.0f}  "
                 f"other={int(m_other.sum()):,}/調後{post[m_other].sum():,.0f}")
        L.append(f"     H分佈: " + " | ".join(
            f"{h}={int((m & hid.eq(h)).sum())}行/{post[m & hid.eq(h)].sum():,.0f}萬"
            for h in hid[m].value_counts().index[:5]))

    # ── 6. pt_class_H=Staff rows: do any have 人工|一級標簽=Staff in 25? ─────────
    L += ["", "=" * 70, "## 6. galaxy 25 staff: pt_class_H value distribution"]
    ph = df.get("項目組H", pd.Series("", index=df.index)).astype(str).str.strip()
    vc = ph[gx25_staff].value_counts().head(15)
    for v, c in vc.items():
        s = post[gx25_staff & ph.eq(v)].sum()
        L.append(f"  {str(v)[:40]:<42} {c:>7,} rows  調後Σ={s:>9,.0f}萬")

    # ── 7. Galaxy 25 staff accounts that DIDN'T exist in 24 (new in 25) ──────────
    L += ["", "=" * 70, "## 7. galaxy 25 staff: new accounts (not in 24 staff)"]
    accts_24 = set(acct[gx24_staff].unique())
    mask_new = gx25_staff & ~acct.isin(accts_24)
    L.append(f"  New accounts (in 25 staff, absent in 24 staff): {int(mask_new.sum()):,} rows  Σ={post[mask_new].sum():,.0f}萬")
    if mask_new.any():
        ag_new = df[mask_new].copy()
        ag_new["_post"] = post[mask_new].values
        ag_new["_acct"] = acct[mask_new].values
        ag_new["_desc"] = adesc[mask_new].values
        s = (ag_new.groupby(["_acct","_desc"])
                   .agg(rows=("_post","count"), post_sum=("_post","sum"))
                   .sort_values("post_sum", ascending=False).head(15))
        for (ac, ad), r in s.iterrows():
            L.append(f"  {str(ac):<12} {str(ad)[:40]:<42} {int(r['rows']):>6,}  {r['post_sum']:>10,.0f}萬")

    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
