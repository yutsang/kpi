"""Diagnose the MGM 25 CAPEX shortfall (ours 188M vs golden 1,097M).

step0_prebuild extracts f1 CAPEX with several filters. This replays them ON f1
step-by-step and prints rows + Σamount dropped at each stage, so we see exactly
which filter eats the ~909M. Prime suspect: the Construction-In-Progress (CIP)
filter, which strips the big construction capex the project-team golden DOES count.

Run (Windows, f1 must be in data/mgm/raw/):
  python scripts/mgm_capex_diagnose.py
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
import pandas as pd
import yaml

AMT = ["Debit minus Credit", "Amount", "Transaction Debit minus Credit"]


def _to_mop(s):
    return pd.to_numeric(s.astype(str).str.replace(",", "").str.strip(), errors="coerce").fillna(0)


def stage(df, amt, label):
    print(f"  {label:<46} rows={len(df):>7,}  Σ={df[amt].sum()/1e6:>12,.1f}M")


def main():
    cfg = yaml.safe_load(Path("conf/company_6/parameters.yml").read_text(encoding="utf-8"))
    src = cfg.get("prebuild_sources") or {}
    f1 = src.get("f1") or ""
    sheet = cfg.get("prebuild_sheet_f1", "2025 Raw Data")
    raw_dir = Path("data/mgm/raw")
    fp = raw_dir / f1
    if not fp.exists():
        print(f"X {fp} not found — confirm f1 name/path"); sys.exit(1)
    print(f"f1 = {f1}\nsheet = {sheet}\n")

    df = pd.read_excel(fp, sheet_name=sheet, dtype=str, engine="openpyxl")
    print(f"columns: {list(df.columns)}\n")
    amt = next((c for c in AMT if c in df.columns), None)
    if not amt:
        print(f"X no amount col among {AMT}"); sys.exit(1)
    df[amt] = _to_mop(df[amt])

    print("=== filter cascade (each line = state AFTER that filter) ===")
    stage(df, amt, "0. raw f1 sheet")

    if "Category" in df.columns:
        d1 = df[df["Category"].fillna("").astype(str).isin(["Gaming", "Non-Gaming"])]
    else:
        d1 = df; print("  (no Category column)")
    stage(d1, amt, "1. Category in [Gaming, Non-Gaming]")

    # offset
    d2 = d1
    off_cols = [c for c in d1.columns if str(c).strip().lower()=="offset" or ("filter" in str(c).lower() and "offset" in str(c).lower())]
    if not off_cols and "Remarks" in d1.columns: off_cols = ["Remarks"]
    for c in off_cols:
        d2 = d2[~d2[c].fillna("").astype(str).str.strip().str.lower().isin({"y","offset"})]
    stage(d2, amt, f"2. offset filter (cols={off_cols})")

    # CIP — the suspect
    l4 = "Ledger Hierarchy Level 4"
    if l4 in d2.columns:
        is_cip = d2[l4].fillna("").astype(str).str.contains("Construction In Progress", na=False)
        d3 = d2[~is_cip]
        print(f"  >>> CIP rows DROPPED: {int(is_cip.sum()):,}  Σ={d2[is_cip][amt].sum()/1e6:,.1f}M  <<<")
    else:
        d3 = d2; print(f"  (no '{l4}' column)")
    stage(d3, amt, "3. after CIP filter (Construction In Progress removed)")

    # reversed
    d4 = d3
    for c in ("Reversed", "Reversed?"):
        if c in d4.columns:
            d4 = d4[d4[c].fillna("").astype(str).str.strip().str.lower() != "yes"]; break
    stage(d4, amt, "4. after reversed filter")

    print(f"\n=== top Ledger L4 in the RAW f1 (where the big capex lives) ===")
    if l4 in df.columns:
        g = df.assign(_l4=df[l4].fillna("").astype(str)).groupby("_l4")[amt].sum().sort_values(key=abs, ascending=False)
        for k, v in g.head(20).items():
            print(f"  {str(k)[:50]:<50} {v/1e6:>12,.1f}M")
    print(f"\nGolden 25 capex = 1,097.3M ; our extract (stage 4) = {d4[amt].sum()/1e6:,.1f}M")


if __name__ == "__main__":
    main()
