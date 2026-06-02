"""Re-derive step 3 summary stats from <company>_unique_signatures.xlsx.

Useful when the original step 3 console output has scrolled away or was lost
because the terminal closed without log redirect.
"""
from __future__ import annotations

import pandas as pd

from kpi.lib.conf import load_config, path
from kpi.lib.io_setup import force_unbuffered_io

force_unbuffered_io()


def main():
    cfg = load_config()
    company = cfg["company"]["code"]
    interim = path(cfg, "interim_dir")
    sig_xlsx = interim / f"{company}_unique_signatures.xlsx"
    if not sig_xlsx.exists():
        raise FileNotFoundError(f"Run step 3 first. Missing: {sig_xlsx}")

    df = pd.read_excel(sig_xlsx)
    n_total = len(df)
    print(f"\n=== Step 3 summary for {company} ===")
    print(f"  total signatures: {n_total:,}")
    print(f"  source xlsx: {sig_xlsx}\n")

    src = df["tag_source"].astype("string").fillna("")
    n_rule       = src.str.startswith("rule:").sum()
    n_acct_code  = src.str.startswith("acct_code:").sum()
    n_manual_acc = (src == "account_manual").sum()
    n_rerank     = src.str.startswith("rerank:").sum()
    n_llm        = (src == "llm").sum()

    status = df.get("llm_status", pd.Series([""] * n_total)).astype("string").fillna("")
    n_needs_manual = (status == "needs_manual").sum()
    n_ok = (status == "ok").sum()

    conf = pd.to_numeric(df.get("llm_confidence"), errors="coerce")
    n_low_conf_llm = ((src == "llm") & (conf < 0.7)).sum()
    n_low_conf_rerank = (src.str.startswith("rerank:") & (conf < 0.5)).sum()

    rt = df.get("row_type", pd.Series([""] * n_total)).astype("string").fillna("")
    n_adjustment = (rt == "adjustment").sum()
    n_normal = (rt == "normal").sum()

    print("=== Tag source breakdown (where each signature got its horizontal_id) ===")
    print(f"  rule (pattern):        {n_rule:>7,}  ({n_rule/n_total*100:5.1f}%)")
    print(f"  rule (account_code):   {n_acct_code:>7,}  ({n_acct_code/n_total*100:5.1f}%)")
    print(f"  account_manual:        {n_manual_acc:>7,}  ({n_manual_acc/n_total*100:5.1f}%)")
    print(f"  rerank:                {n_rerank:>7,}  ({n_rerank/n_total*100:5.1f}%)")
    print(f"  llm:                   {n_llm:>7,}  ({n_llm/n_total*100:5.1f}%)")
    print(f"  needs_manual (no tag): {n_needs_manual:>7,}  ({n_needs_manual/n_total*100:5.1f}%)")

    print("\n=== Quality flags ===")
    print(f"  low-confidence LLM (<0.7):    {n_low_conf_llm:>5,}")
    print(f"  low-confidence rerank (<0.5): {n_low_conf_rerank:>5,}")
    print(f"  row_type=adjustment:          {n_adjustment:>5,}  (auto-flagged via has_negative)")
    print(f"  row_type=normal:              {n_normal:>5,}")

    print("\n=== Horizontal distribution (by signature count) ===")
    h_label = df.get("llm_horizontal_label", pd.Series([""] * n_total)).astype("string").fillna("(未分類)")
    h_label = h_label.where(h_label.ne(""), "(未分類)")
    print(h_label.value_counts().to_string())

    print("\n=== Total amount by horizontal ===")
    if "total_amount" in df.columns:
        amt = pd.to_numeric(df["total_amount"], errors="coerce").fillna(0)
        agg = pd.DataFrame({"label": h_label, "amount": amt})
        print(agg.groupby("label")["amount"].agg(["count", "sum"]).sort_values("sum", ascending=False).to_string())

    print(f"\n=== Top 10 'needs_manual' signatures by amount ===")
    if n_needs_manual > 0:
        nm = df[status == "needs_manual"].copy()
        nm["abs_amt"] = pd.to_numeric(nm["total_amount"], errors="coerce").abs()
        cols = ["account_code", "account_desc", "desc_norm", "row_count", "total_amount", "llm_error"]
        cols = [c for c in cols if c in nm.columns]
        print(nm.sort_values("abs_amt", ascending=False).head(10)[cols].to_string(index=False))
