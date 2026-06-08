"""MGM 2025 — verify the project-team highlighted H corrections before wiring conf.

The project team flagged (for MGM 25):
  WD1 by Ledger Account:  430120→餐飲 463120→Comp其他 420120→酒店客房 463140→活動場地
       463130→Comp其他 463150→贈票 430140→活動場地 410140→餐飲 412140→餐飲  (剩下→其他)
  PM  by Item Type:       Room/Hotel Front Desk/Mandarin Oriental→酒店客房  Food & Beverage→餐飲
       Other/Vouchers→Comp其他  (剩下→其他)
  CAPEX by Supplier:      3812 GALLERY / POLY AUCTION / BONHAMS / SOTHEBY'S → 藝術品 (lock 博物館?)
  540025 Property Promo breakdown (WD1):
       JLM/Spend Cat 'digital content'→其他   Supplier 保利/poly→其他
       Supplier 第十五屆全國運動會…澳門賽區籌備辦公室 / 體育基金 / Fundo de Turismo→贊助
       JLM/Spend Cat 'Artist Performance Fee'→專業服務費

This dumps, for MGM 25 rows, everything needed to size + wire those rules: column presence,
Source distinct, per-account current H, Item Type × Source, art vendors, 540025 breakdown,
and crucially the H_Label pre-set coverage on the affected rows (decides rule ORДЕР).

Run (Windows):  python scripts/inspect_mgm_25_comments.py
Output: prints + results/inspect_mgm_25_comments.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
TR = ROOT / "data" / "mgm" / "interim" / "company_6_tagged_rows.parquet"
AMT = "Debit minus Credit"
AC = "Ledger Account"
SRC = "Source"
SUP = "Supplier"
JLM = "Journal Line Memo"
ITEM = "Item Type"
SPENDCAT = "Spend Category as Worktag"
HL = "H_Label"
COMP_ACCTS = ["430120", "463120", "420120", "463140", "463130", "463150", "430140", "410140", "412140"]
ART_VENDORS = ["3812 GALLERY", "POLY AUCTION", "BONHAMS", "SOTHEBY"]
PROMO_SUP = ["保利", "poly", "第十五屆全國運動會", "澳門賽區籌備辦公室", "體育基金", "Fundo de Turismo"]
PROMO_KW = ["digital content", "Artist Performance Fee"]


def main():
    L = ["# inspect_mgm_25_comments — verify project-team H corrections before wiring"]
    if not TR.exists():
        L.append(f"X {TR} missing — run kedro mgm first"); _w(L); return
    df = pd.read_parquet(TR)
    per = next((c for c in ("report_period", "report_year", "years") if c in df.columns), None)
    if per:
        df = df[df[per].astype(str).str.startswith("25")].copy()
    amt = AMT if AMT in df.columns else next((c for c in df.columns if "Debit" in str(c) or "Amount" in str(c)), None)
    a = pd.to_numeric(df[amt], errors="coerce").fillna(0.0)
    tot = a.abs().sum() or 1

    def S(col):
        return df[col].astype("string").fillna("").str.strip() if col in df.columns else None

    L.append(f"\n25 rows={len(df):,}  Σ={a.sum():,.0f}  Σ|amt|={a.abs().sum():,.0f}  amount={amt!r}")
    # (0) column presence
    L.append("\n## (0) column presence:")
    for c in (AC, SRC, SUP, JLM, ITEM, SPENDCAT, HL, "horizontal_label", "vertical_label", "final_capex_opex"):
        L.append(f"   {c:32s} {'PRESENT' if c in df.columns else '--- MISSING ---'}")

    # (1) Source distinct
    src = S(SRC)
    if src is not None:
        L.append(f"\n## (1) Source distinct (數據源)  Σ|amt| M:")
        for v, x in a.abs().groupby(src.replace('', '(blank)')).sum().sort_values(ascending=False).items():
            L.append(f"   {str(v)[:24]:24s} {x/1e6:9.2f}M  {x/tot*100:5.1f}%")

    hl = S(HL)
    hlab = S("horizontal_label")

    # (2) WD1 comp accounts — current H + H_Label + amount
    ac = S(AC)
    if ac is not None:
        L.append(f"\n## (2) WD1 comp accounts — Ledger Account → current H / H_Label / Σ|amt|:")
        for code in COMP_ACCTS:
            m = ac.str.startswith(code)
            if not m.any():
                L.append(f"   {code}: (none)"); continue
            cur = hlab[m].replace('', '(blank)').value_counts().head(3).to_dict() if hlab is not None else {}
            hlv = hl[m].replace('', '(blank)').value_counts().head(3).to_dict() if hl is not None else {}
            L.append(f"   {code}: {int(m.sum()):,} rows  {a.abs()[m].sum()/1e6:7.2f}M  curH={cur}  H_Label={hlv}")
        # full Ledger Account distribution within WD1 (to size '剩下→其他')
        if src is not None:
            w1 = src.eq("WD1")
            L.append(f"\n## (2b) ALL Ledger Accounts within Source=WD1 ({int(w1.sum()):,} rows, {a.abs()[w1].sum()/1e6:.1f}M)  Σ|amt| M:")
            g = a.abs()[w1].groupby(ac[w1].replace('', '(blank)')).sum().sort_values(ascending=False)
            for v, x in g.head(30).items():
                L.append(f"   {str(v)[:30]:30s} {x/1e6:8.2f}M")

    # (3) Item Type (PM) — distinct + which Source + current H
    it = S(ITEM)
    if it is not None and src is not None:
        L.append(f"\n## (3) Item Type distinct (Σ|amt| M) + dominant Source + current H:")
        for v, x in a.abs().groupby(it.replace('', '(blank)')).sum().sort_values(ascending=False).items():
            if str(v) == '(blank)':
                continue
            mm = it.eq(v)
            srcs = src[mm].replace('', '(blank)').value_counts().head(2).to_dict()
            cur = hlab[mm].replace('', '(blank)').value_counts().head(2).to_dict() if hlab is not None else {}
            L.append(f"   {str(v)[:24]:24s} {x/1e6:8.2f}M  src={srcs}  curH={cur}")
        L.append(f"   (Item Type non-blank rows: {int(it.ne('').sum()):,}; blank: {int(it.eq('').sum()):,})")

    # (4) CAPEX art vendors
    sup = S(SUP)
    co = S("final_capex_opex")
    if sup is not None:
        L.append(f"\n## (4) Art vendors (any Source) — Supplier match → current H / capex / Σ|amt| / V:")
        for nm in ART_VENDORS:
            m = sup.str.contains(nm, case=False, na=False)
            if not m.any():
                L.append(f"   {nm}: (none)"); continue
            cur = hlab[m].replace('', '(blank)').value_counts().head(2).to_dict() if hlab is not None else {}
            cox = co[m].value_counts().head(2).to_dict() if co is not None else {}
            vv = S("vertical_label"); vd = vv[m].replace('', '(blank)').value_counts().head(2).to_dict() if vv is not None else {}
            L.append(f"   {nm:18s} {int(m.sum()):,} rows  {a.abs()[m].sum()/1e6:7.3f}M  curH={cur}  capex={cox}  V={vd}")

    # (5) 540025 breakdown
    if ac is not None:
        m025 = ac.str.startswith("540025")
        L.append(f"\n## (5) 540025 Property Promotional — {int(m025.sum()):,} rows  {a.abs()[m025].sum()/1e6:.2f}M  curH:")
        if hlab is not None:
            for v, x in a.abs()[m025].groupby(hlab[m025].replace('', '(blank)')).sum().sort_values(ascending=False).items():
                L.append(f"     curH {str(v)[:20]:20s} {x/1e6:7.2f}M")
        sub = df[m025]; asub = a.abs()[m025]
        for col, kws in ((SUP, PROMO_SUP), (JLM, PROMO_KW), (SPENDCAT, PROMO_KW)):
            s = sub[col].astype("string").fillna("").str.strip() if col in sub.columns else None
            if s is None:
                L.append(f"   [{col}] MISSING"); continue
            L.append(f"   [{col}] keyword hits within 540025:")
            for kw in kws:
                mm = s.str.contains(kw, case=False, na=False)
                L.append(f"      {kw:46s} {int(mm.sum()):,} rows  {asub[mm].sum()/1e6:7.3f}M")
            L.append(f"      -- top {col} values in 540025 --")
            for v, n in s[s.ne('')].value_counts().head(12).items():
                L.append(f"        {str(v)[:50]:50s} {n}")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_mgm_25_comments.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
