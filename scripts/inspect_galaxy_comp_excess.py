"""inspect_galaxy_comp_excess.py — 診斷 galaxy 25 comp 超額 +31,967萬
比較 25-prefix 三個 bucket 的 comp 分佈 + account code breakdown。
Run: python scripts\inspect_galaxy_comp_excess.py
Out: results\inspect_galaxy_comp_excess.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV  = ROOT / "tableau_combined_25.csv"
OUT  = ROOT / "results" / "inspect_galaxy_comp_excess.txt"
COMP_H = {"H_HOTEL_ROOM", "H_VENUE", "H_FNB", "H_COMP_TICKET", "H_COMP_OTHER"}


def main():
    L = ["# inspect_galaxy_comp_excess — galaxy 25 comp +31,967 診斷", ""]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return

    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    amt = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    yb  = df["year_bucket"].astype(str)
    y2  = yb.str[:2]
    hid = df.get("horizontal_id", pd.Series("", index=df.index)).astype(str).str.strip()
    ent = df["entity"].astype(str)
    co  = df.get("final_capex_opex", pd.Series("", index=df.index)).astype(str).str.strip()
    acct = df.get("account_code", pd.Series("", index=df.index)).astype(str).str.strip()
    adesc = df.get("account_desc", pd.Series("", index=df.index)).astype(str).str.strip()
    ph   = df.get("項目組H", pd.Series("", index=df.index)).astype(str).str.strip()

    gx = ent.eq("galaxy")
    gx25 = gx & y2.eq("25")
    comp25 = gx25 & hid.isin(COMP_H)

    # ── 1. galaxy 25 comp by bucket ─────────────────────────────────────────────
    L += ["=" * 70, "## 1. galaxy 25 comp by year_bucket (golden=59,828)"]
    total = 0.0
    for bk in ["25", "25_24SY", "25_23SY"]:
        m = comp25 & yb.eq(bk)
        s = amt[m].sum()
        total += s
        rows = int(m.sum())
        L.append(f"  {bk:<12}: {rows:>8,} rows  Σ={s:>10,.0f}萬")
    L.append(f"  {'TOTAL':<12}: {int(comp25.sum()):>8,} rows  Σ={total:>10,.0f}萬  (golden 59,828  Δ={total-59828:+,.0f})")

    # ── 2. comp by horizontal_id + bucket ────────────────────────────────────────
    L += ["", "=" * 70, "## 2. comp by H bucket × year_bucket"]
    for h in sorted(COMP_H):
        m = comp25 & hid.eq(h)
        total_h = amt[m].sum()
        row_h   = int(m.sum())
        bkt_str = " | ".join(
            f"{bk}:{int((m & yb.eq(bk)).sum())} rows/{amt[m & yb.eq(bk)].sum():,.0f}萬"
            for bk in ["25", "25_24SY", "25_23SY"] if int((m & yb.eq(bk)).sum()) > 0
        )
        L.append(f"  {h:<20} total={row_h:>7,}/{total_h:>9,.0f}萬  ← {bkt_str}")

    # ── 3. top account codes in galaxy 25 comp ──────────────────────────────────
    L += ["", "=" * 70, "## 3. top account codes in galaxy 25 comp (by Σ)"]
    ag = df[comp25].copy()
    ag["amt"] = amt[comp25].values
    summary = (ag.groupby(["account_code", "account_desc", "horizontal_id"])
                 .agg(rows=("amt", "count"), total=("amt", "sum"))
                 .sort_values("total", ascending=False)
                 .head(30))
    for (ac, ad, hh), row in summary.iterrows():
        L.append(f"  {str(ac):<15} {str(ad)[:38]:<40} {str(hh):<20} rows={int(row['rows']):>6,}  Σ={row['total']:>9,.0f}萬")

    # ── 4. compare 25 bucket vs 25_24SY+25_23SY ─────────────────────────────────
    L += ["", "=" * 70, "## 4. galaxy 25 comp: 純25 vs carry-forward (25_24SY+25_23SY)"]
    pure25 = comp25 & yb.eq("25")
    carry  = comp25 & yb.isin(["25_24SY", "25_23SY"])
    L.append(f"  純25 bucket:          {int(pure25.sum()):>8,} rows  Σ={amt[pure25].sum():>10,.0f}萬")
    L.append(f"  25_24SY+25_23SY:      {int(carry.sum()):>8,} rows  Σ={amt[carry].sum():>10,.0f}萬")
    L.append(f"  ─ 如 golden 只計純25：Δ={amt[pure25].sum()-59828:+,.0f}萬")
    L.append(f"  ─ 如 golden 計全部：  Δ={amt[comp25].sum()-59828:+,.0f}萬")

    # ── 5. galaxy 24 comp by bucket (for comparison) ────────────────────────────
    L += ["", "=" * 70, "## 5. galaxy 24 comp by bucket (golden=55,321, ref ✓)"]
    gx24 = gx & y2.eq("24")
    comp24 = gx24 & hid.isin(COMP_H)
    total24 = 0.0
    for bk in ["24", "24_23SY"]:
        m = comp24 & yb.eq(bk)
        s = amt[m].sum()
        total24 += s
        L.append(f"  {bk:<12}: {int(m.sum()):>8,} rows  Σ={s:>10,.0f}萬")
    L.append(f"  {'TOTAL':<12}: {int(comp24.sum()):>8,} rows  Σ={total24:>10,.0f}萬  (golden 55,321  Δ={total24-55321:+,.0f})")

    # ── 6. account codes in 25 comp NOT heavily in 24 comp ───────────────────────
    L += ["", "=" * 70, "## 6. account codes in 25-prefix comp but small/absent in 24-prefix comp"]
    c24 = df[comp24].copy(); c24["amt"] = amt[comp24].values
    c25 = df[comp25].copy(); c25["amt"] = amt[comp25].values
    sum24 = c24.groupby("account_code")["amt"].sum().rename("sum24")
    sum25 = c25.groupby("account_code")["amt"].sum().rename("sum25")
    cmp = pd.concat([sum25, sum24], axis=1).fillna(0)
    cmp["excess"] = cmp["sum25"] - cmp["sum24"]
    cmp["ratio"] = cmp["sum25"] / (cmp["sum24"] + 0.01)
    cmp = cmp.sort_values("excess", ascending=False).head(20)
    L.append("  account_code    Σ_25       Σ_24       excess     ratio")
    for ac, row in cmp.iterrows():
        desc = adesc[acct.eq(str(ac))].mode()
        desc_str = str(desc.iloc[0])[:35] if len(desc) else ""
        L.append(f"  {str(ac):<15} {row['sum25']:>9,.0f}  {row['sum24']:>9,.0f}  {row['excess']:>9,.0f}  {row['ratio']:>6.1f}x  {desc_str}")

    # ── 7. 項目組H distribution in galaxy 25 comp ───────────────────────────────
    L += ["", "=" * 70, "## 7. 項目組H values in galaxy 25 comp (top20)"]
    vc = ph[comp25].value_counts().head(20)
    for v, c in vc.items():
        s = amt[comp25 & ph.eq(v)].sum()
        L.append(f"  {str(v)[:40]:<42} {c:>8,} rows  Σ={s:>9,.0f}萬")

    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
