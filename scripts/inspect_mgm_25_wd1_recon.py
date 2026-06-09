"""MGM 2025 — is combine's WD1 gross-of-reversals problem systemic, or only 430120?
combine 430120 = 23.8M (614 rows, almost all +) but NGO_WD1 430120 = 12.59M net (775 rows).
Hypothesis: combine dropped the credit/reversal rows → overstates comp ~2×. This compares
combine(Source=WD1) vs NGO_WD1 per Ledger Account (net / pos / neg / rows) and totals the gap,
so we see how much of MGM 2025 is inflated.

Run (Windows):  python scripts/inspect_mgm_25_wd1_recon.py
Output: prints + results/inspect_mgm_25_wd1_recon.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "data" / "mgm" / "raw" / "mgm_25_raw.xlsx"
ACC, SRC, AMT = "Ledger Account", "Source", "Debit minus Credit"


def _col(df, *names):
    for name in names:
        if name in df.columns: return name
        for c in df.columns:
            if str(c).strip() == name: return c
    return None


def _agg(df, ac, am):
    amt = pd.to_numeric(df[am], errors="coerce").fillna(0.0)
    k = df[ac].astype("string").fillna("(blank)")
    g = pd.DataFrame({"k": k, "a": amt}).groupby("k")
    return pd.DataFrame({
        "net": g["a"].sum(), "pos": g["a"].apply(lambda x: x[x > 0].sum()),
        "neg": g["a"].apply(lambda x: x[x < 0].sum()), "n": g["a"].size(),
    })


def main():
    L = ["# inspect_mgm_25_wd1_recon — combine(WD1) vs NGO_WD1 per Ledger Account (gross-vs-net)"]
    if not RAW.exists():
        L.append(f"X {RAW} missing"); _w(L); return
    cmb = pd.read_excel(RAW, sheet_name="combine")
    ngo = pd.read_excel(RAW, sheet_name="NGO_WD1")
    ca, cm, cs = _col(cmb, ACC), _col(cmb, AMT), _col(cmb, SRC)
    na, nm = _col(ngo, ACC), _col(ngo, AMT)
    cmb_wd1 = cmb[cmb[cs].astype("string").fillna("").str.strip().eq("WD1")] if cs else cmb
    C = _agg(cmb_wd1, ca, cm)
    N = _agg(ngo, na, nm)
    L.append(f"\nTOTAL  combine(WD1): net={C['net'].sum()/1e6:,.1f}M pos={C['pos'].sum()/1e6:,.1f}M "
             f"neg={C['neg'].sum()/1e6:,.1f}M rows={int(C['n'].sum()):,}")
    L.append(f"TOTAL  NGO_WD1     : net={N['net'].sum()/1e6:,.1f}M pos={N['pos'].sum()/1e6:,.1f}M "
             f"neg={N['neg'].sum()/1e6:,.1f}M rows={int(N['n'].sum()):,}")
    L.append(f"TOTAL  GAP (combine−NGO net): {(C['net'].sum()-N['net'].sum())/1e6:,.1f}M  "
             f"← if large+positive, combine systematically dropped reversals")

    j = C.join(N, how="outer", lsuffix="_cmb", rsuffix="_ngo").fillna(0.0)
    j["gap"] = j["net_cmb"] - j["net_ngo"]
    j = j.reindex(j["gap"].abs().sort_values(ascending=False).index)
    L.append("\n## top Ledger Accounts by |combine−NGO net gap| (M):")
    L.append(f"   {'account':40s} {'cmb_net':>9s} {'ngo_net':>9s} {'gap':>9s}  {'cmb_n':>6s}{'ngo_n':>6s}")
    for k, r in j.head(30).iterrows():
        L.append(f"   {str(k)[:40]:40s} {r['net_cmb']/1e6:>9.2f} {r['net_ngo']/1e6:>9.2f} "
                 f"{r['gap']/1e6:>9.2f}  {int(r['n_cmb']):>6}{int(r['n_ngo']):>6}")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_mgm_25_wd1_recon.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
