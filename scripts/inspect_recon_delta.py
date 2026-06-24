"""inspect_recon_delta.py — 診斷 recon delta 根因 (galaxy comp/staff 25 + sjm capex 25 + galaxy capex 25)
Run: python scripts\inspect_recon_delta.py
Out: results\inspect_recon_delta.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "inspect_recon_delta.txt"
COMP_H = {"H_HOTEL_ROOM", "H_VENUE", "H_FNB", "H_COMP_TICKET", "H_COMP_OTHER"}


def main():
    L = ["# inspect_recon_delta — comp/staff/capex 差根因（galaxy/sjm 25年）", ""]
    if not CSV.exists():
        L.append("!! tableau_combined_25.csv 揾唔到"); _w(L); return

    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    post = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    yb  = df["year_bucket"].astype(str)
    y2  = yb.str[:2]
    co  = df.get("final_capex_opex", pd.Series("", index=df.index)).astype(str).str.strip()
    hid = df.get("horizontal_id", pd.Series("", index=df.index)).astype(str).str.strip()
    ent = df["entity"].astype(str)
    pt  = df.get("項目組H", pd.Series("", index=df.index)).astype(str)
    has_ptH = pt.str.strip().ne("")

    # ── 1. galaxy 25 comp ──────────────────────────────────────────────────────
    L += ["=" * 70, "## 1. galaxy 25 COMP  (golden=59,828)"]
    gx25 = (ent == "galaxy") & (y2 == "25")
    gx25_comp = gx25 & hid.isin(COMP_H)
    L.append(f"  total 25 rows: {gx25.sum():,}   comp rows: {gx25_comp.sum():,}   Σ調整後={post[gx25_comp].sum():,.0f}萬")
    # pt_class_H coverage among comp rows
    L.append(f"  comp WITH 項目組H:    {(gx25_comp & has_ptH).sum():,}   Σ={post[gx25_comp & has_ptH].sum():,.0f}萬")
    L.append(f"  comp WITHOUT 項目組H: {(gx25_comp & ~has_ptH).sum():,}   Σ={post[gx25_comp & ~has_ptH].sum():,.0f}萬")
    # by bucket
    L.append("  by bucket:")
    for bk in ["25", "25_24SY", "25_23SY"]:
        m = gx25_comp & (yb == bk)
        L.append(f"    {bk:<12}: {m.sum():>6,} rows  Σ={post[m].sum():>10,.0f}萬")
    # 項目組H top values for galaxy 25
    L.append("  項目組H value distribution (galaxy 25 ALL rows — top15):")
    vc = pt[gx25].value_counts(dropna=False).head(15)
    for v, c in vc.items():
        L.append(f"    {repr(str(v)):<40} {c:>8,}")

    # ── 2. galaxy 25 staff ─────────────────────────────────────────────────────
    L += ["", "=" * 70, "## 2. galaxy 25 STAFF opex  (golden=14,568)"]
    gx25_staff = gx25 & hid.eq("H_LABOR") & (~co.eq("Capex"))
    L.append(f"  staff rows: {gx25_staff.sum():,}   Σ調整後={post[gx25_staff].sum():,.0f}萬")
    L.append("  by bucket:")
    for bk in ["25", "25_24SY", "25_23SY"]:
        m = gx25_staff & (yb == bk)
        L.append(f"    {bk:<12}: {m.sum():>6,} rows  Σ={post[m].sum():>10,.0f}萬")
    # account_code top for galaxy 25 staff
    if "account_code" in df.columns:
        L.append("  top account_code for galaxy 25 staff (top10):")
        vc2 = df.loc[gx25_staff, "account_code"].value_counts().head(10)
        for v, c in vc2.items():
            L.append(f"    {str(v):<20} rows={c:,}  Σ={post[gx25_staff & (df['account_code']==v)].sum():,.0f}萬")

    # ── 3. sjm 25 capex by bucket ──────────────────────────────────────────────
    L += ["", "=" * 70, "## 3. sjm 25 CAPEX by bucket  (golden=117,736)"]
    for bk in ["25", "25_24SY", "25_23SY"]:
        m_all = (ent == "sjm") & (yb == bk)
        m_cap = m_all & co.eq("Capex")
        m_opx = m_all & co.eq("Opex")
        L.append(f"  {bk:<12}: total={m_all.sum():>6,}  capex={m_cap.sum():>6,} Σ={post[m_cap].sum():>9,.0f}  opex={m_opx.sum():>6,} Σ={post[m_opx].sum():>9,.0f}")
    sj25_cap = (ent=="sjm") & (y2=="25") & co.eq("Capex")
    L.append(f"  SUM 25-prefix capex = {post[sj25_cap].sum():,.0f}萬  (golden 117,736 →Δ={post[sj25_cap].sum()-117736:+,.0f})")

    # ── 4. galaxy 25 capex ─────────────────────────────────────────────────────
    L += ["", "=" * 70, "## 4. galaxy 25 CAPEX by bucket  (golden=122,929)"]
    for bk in ["25", "25_24SY", "25_23SY"]:
        m_all = (ent == "galaxy") & (yb == bk)
        m_cap = m_all & co.eq("Capex")
        L.append(f"  {bk:<12}: capex={m_cap.sum():>5,} rows  Σ={post[m_cap].sum():>9,.0f}萬  (total rows={m_all.sum():,})")
    gx25_cap = gx25 & co.eq("Capex")
    L.append(f"  SUM 25-prefix capex = {post[gx25_cap].sum():,.0f}萬  (golden 122,929 →Δ={post[gx25_cap].sum()-122929:+,.0f})")

    # ── 5. 全家 summary ────────────────────────────────────────────────────────
    L += ["", "=" * 70, "## 5. 全家 25年 comp/staff/capex summary"]
    goldens = {
        "comp":  {"galaxy":59828,"wynn":19823,"vml":8396,"melco":8403,"mgm":7532,"sjm":2920},
        "staff": {"galaxy":14568,"wynn":9923,"vml":809,"melco":27742,"mgm":5813,"sjm":142},
        "capex": {"galaxy":122929,"wynn":139186,"vml":134692,"melco":106994,"mgm":63272,"sjm":117736},
    }
    hdr = f"  {'entity':<8}{'capex':>10}{'g_cap':>10}{'Δcap':>8}{'comp':>10}{'g_comp':>10}{'Δcomp':>8}{'staff':>10}{'g_stf':>10}{'Δstf':>8}"
    L.append(hdr)
    for e in ["galaxy","wynn","vml","melco","mgm","sjm"]:
        m25 = (ent==e) & (y2=="25")
        cap  = post[m25 & co.eq("Capex")].sum()
        stf  = post[m25 & hid.eq("H_LABOR") & ~co.eq("Capex")].sum()
        cmp  = post[m25 & hid.isin(COMP_H)].sum()
        gcap = goldens["capex"][e]; gstf = goldens["staff"][e]; gcmp = goldens["comp"][e]
        L.append(f"  {e:<8}{cap:>10,.0f}{gcap:>10,}{cap-gcap:>+8,.0f}{cmp:>10,.0f}{gcmp:>10,}{cmp-gcmp:>+8,.0f}{stf:>10,.0f}{gstf:>10,}{stf-gstf:>+8,.0f}")

    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
