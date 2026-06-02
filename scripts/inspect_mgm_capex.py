"""MGM capex tie-out — straight from f1 '2025 Raw Data' (the CAPEX JE detail).

The 257M gap = the 'Construction In Progress(L4)' line (−258M) PLUS multi-year Budget Sources.
This finds which Budget Source × L4 scope reproduces the golden capex (Σ ~1.097B), and dumps
capex per project (Project Plan Task) vs golden so we can tie by project.

Run on the machine with data/mgm/raw/:
  python scripts/inspect_mgm_capex.py
"""
from __future__ import annotations
import argparse, re, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parent.parent


def _golden_capex():
    p = ROOT / "results" / "mgm_golden_25.tsv"
    if not p.exists():
        return None
    tot = 0.0
    for line in p.read_text(encoding="utf-8-sig").splitlines():
        c = [x.strip() for x in line.split("\t")]
        if len(c) >= 6 and re.fullmatch(r"\d+", c[0]):
            try:
                tot += float(c[3].replace(",", "")) * 10000
            except Exception:
                pass
    return tot


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--raw", default="data/mgm/raw")
    args = p.parse_args()
    cfg = yaml.safe_load((ROOT / "conf/company_6/parameters.yml").read_text(encoding="utf-8"))
    f1 = (cfg.get("prebuild_sources") or {}).get("f1")
    fp = ROOT / args.raw / f1
    if not fp.exists():
        print(f"X {fp} missing"); return
    df = pd.read_excel(fp, sheet_name="2025 Raw Data", header=0, dtype=str)
    amt = pd.to_numeric(df.get("Debit minus Credit"), errors="coerce").fillna(0)
    bs = df.get("Budget Source", pd.Series([""] * len(df))).astype(str).str.strip()
    l4 = df.get("Ledger Hierarchy Level 4", pd.Series([""] * len(df))).astype(str).str.strip()
    g = _golden_capex()
    print(f"f1 '2025 Raw Data': {len(df):,} rows, Σ={amt.sum():,.0f}   golden capex≈{g:,.0f}" if g else "")

    print("\n=== Budget Source × L4 (是邊個 scope = golden?) ===")
    cross = amt.groupby([bs, l4]).sum().reset_index(name="amt")
    cross = cross.reindex(cross["amt"].abs().sort_values(ascending=False).index)
    for _, r in cross.head(25).iterrows():
        print(f"  {str(r[0])[:24]:24} | {str(r[1])[:36]:36} {r['amt']:>16,.0f}")

    print("\n=== 累積 scenario vs golden ===")
    yr = bs.str.extract(r"(\d{4})")[0]
    for keep in (["2025"], ["2024", "2025"], ["2023", "2024", "2025"]):
        m = yr.isin(keep)
        for cip in ("incl CIP", "excl CIP"):
            mm = m & (~l4.str.contains("Construction In Progress", na=False)) if cip == "excl CIP" else m
            s = amt[mm].sum()
            d = (s - g) if g else 0
            print(f"  Budget {keep} / {cip:9}: {s:>16,.0f}   Δgolden={d:>16,.0f}")

    # per-project (Project Plan Task) capex
    pcol = next((c for c in ("Project Plan Task", "Business Entity", "Campaign") if c in df.columns), None)
    if pcol:
        pj = amt.groupby(df[pcol].astype(str).str.strip()).sum().reset_index(name="capex")
        pj = pj.reindex(pj["capex"].abs().sort_values(ascending=False).index)
        out = ROOT / "results" / "mgm_capex_byproject.tsv"
        pj.to_csv(out, sep="\t", index=False, encoding="utf-8-sig")
        print(f"\nper-project capex ({pcol!r}) → {out}  (top:)")
        for _, r in pj.head(15).iterrows():
            print(f"  {str(r[pcol])[:46]:46} {r['capex']:>16,.0f}")


if __name__ == "__main__":
    main()
