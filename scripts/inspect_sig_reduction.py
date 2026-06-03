"""Diagnose why an entity has so many unique signatures + which lever shrinks the
set BEFORE dumping for H classification. Pure counts — nothing leaves.

signature = account_code | account_desc | desc_norm [| job_code]

For data/{ent}/interim/{com}_unique_signatures.xlsx it reports:
  1. current sig count + row/amount coverage
  2. COLLAPSE counts if we key on fewer parts:
        drop job_code      → account|account_desc|desc_norm
        + drop desc_norm   → account|account_desc
        account_code only
     (the drop from each tells you how much that dimension is over-splitting)
  3. job_code cardinality (the 4th dimension, if present)
  4. NORMALIZER GAPS: sigs whose desc_norm STILL contains digits/long tokens
     — time-series / IDs the normalizer missed → candidates for new strip rules.
     Top 25 shown so you can eyeball the offending format (e.g. 'rent 1月').
  5. AMOUNT COVERAGE: how many top-|amount| sigs cover 90 / 95 / 99 / 99.9% —
     tells you the top-$ dump cutoff (classify those, default the long tail).

Run (Windows):
  python scripts/inspect_sig_reduction.py --entity vml
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
ENTITIES = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
            "vml": "company_4", "melco": "company_5", "mgm": "company_6"}


def _load(ent: str, com: str) -> pd.DataFrame | None:
    cand = [ROOT / "data" / ent / "interim" / f"{com}_unique_signatures.xlsx",
            ROOT / "data" / "review" / "_dump" / f"{ent}_sigs.tsv"]
    for p in cand:
        if p.exists():
            print(f"[{ent}] reading {p.relative_to(ROOT)}")
            return pd.read_excel(p) if p.suffix == ".xlsx" else pd.read_csv(p, sep="\t")
    print(f"[{ent}] X no sig file found ({[str(c.relative_to(ROOT)) for c in cand]})")
    return None


def _job_code(df: pd.DataFrame) -> pd.Series:
    """Recover the job_code tail (4th signature field) robustly via the prefix."""
    if not {"account_code", "account_desc", "desc_norm", "signature"} <= set(df.columns):
        return pd.Series([""] * len(df))
    pref = (df["account_code"].astype(str) + "|" + df["account_desc"].astype(str)
            + "|" + df["desc_norm"].astype(str) + "|")
    sig = df["signature"].astype(str)
    out = []
    for s, p in zip(sig, pref):
        out.append(s[len(p):] if s.startswith(p) else "")
    return pd.Series(out, index=df.index)


def inspect(ent: str):
    com = ENTITIES[ent]
    df = _load(ent, com)
    if df is None:
        return
    n = len(df)
    rc = pd.to_numeric(df.get("row_count", df.get("n_rows", 0)), errors="coerce").fillna(0)
    amt = pd.to_numeric(df.get("total_amount", df.get("amount", 0)), errors="coerce").fillna(0)
    tot_rows, tot_amt = rc.sum(), amt.abs().sum()
    print(f"\n===== {ent}: {n:,} signatures  ({int(tot_rows):,} rows, |Σamt|={tot_amt:,.0f}) =====")

    for c in ("account_code", "account_desc", "desc_norm"):
        if c not in df.columns:
            print(f"  ! missing column {c!r} — collapse report skipped");
    if {"account_code", "account_desc", "desc_norm"} <= set(df.columns):
        ac, ad, dn = df["account_code"].astype(str), df["account_desc"].astype(str), df["desc_norm"].astype(str)
        c_nojob = (ac + "|" + ad + "|" + dn).nunique()
        c_ad = (ac + "|" + ad).nunique()
        c_a = ac.nunique()
        print("\n  COLLAPSE potential (distinct keys):")
        print(f"    current (incl. job_code)        {n:,}")
        print(f"    drop job_code                   {c_nojob:,}   (-{n - c_nojob:,})")
        print(f"    + drop desc_norm (acct+adesc)   {c_ad:,}   (-{c_nojob - c_ad:,})")
        print(f"    account_code only               {c_a:,}")

        jc = _job_code(df)
        nonempty = int((jc.str.strip() != "").sum())
        print(f"\n  job_code 4th-dimension: {jc[jc.str.strip() != ''].nunique():,} distinct, "
              f"on {nonempty:,}/{n:,} sigs ({nonempty / n * 100:.0f}%)")

        # normalizer gaps
        has_digit = dn.str.contains(r"\d", na=False)
        g = df[has_digit]
        gamt = amt[has_digit].abs()
        print(f"\n  NORMALIZER GAPS: {int(has_digit.sum()):,} sigs still carry digits in desc_norm "
              f"(|Σamt|={gamt.sum():,.0f}, {gamt.sum() / tot_amt * 100:.0f}% of $) — top 25 by |amt|:")
        gg = (pd.DataFrame({"desc_norm": dn[has_digit], "rc": rc[has_digit], "amt": amt[has_digit]})
              .groupby("desc_norm").agg(n=("rc", "size"), rows=("rc", "sum"), amt=("amt", "sum")))
        gg = gg.reindex(gg["amt"].abs().sort_values(ascending=False).index)
        for v, r in gg.head(25).iterrows():
            print(f"     {str(v)[:60]:62s} n_sig={int(r['n']):4d} rows={int(r['rows']):6d} Σ={r['amt']:>15,.0f}")

    # amount coverage
    s = amt.abs().sort_values(ascending=False).reset_index(drop=True)
    cum = s.cumsum()
    print("\n  AMOUNT COVERAGE (top-$ dump cutoff):")
    for pct in (0.90, 0.95, 0.99, 0.999):
        k = int((cum < pct * tot_amt).sum()) + 1
        print(f"    top {k:,} sigs cover {pct*100:.1f}% of |Σamt|")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", nargs="+", required=True, choices=sorted(ENTITIES))
    for e in ap.parse_args().entity:
        inspect(e)


if __name__ == "__main__":
    main()
