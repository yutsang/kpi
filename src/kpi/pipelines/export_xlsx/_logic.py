"""Export <company>_tagged_rows.parquet -> xlsx for manual review.

Reads step 4 output and writes all rows (or top-N) sorted by abs(amount) desc.

Optional env vars:
    EXPORT_TOP_ONLY=N   only top N rows by |amount| (faster, smaller file)
    EXPORT_CSV=1        write csv instead of xlsx (5x faster, 3x smaller)
"""
from __future__ import annotations

import os

import pandas as pd

from kpi.lib.conf import load_config, path
from kpi.lib.io_setup import force_unbuffered_io

force_unbuffered_io()


def main():
    cfg = load_config()
    company = cfg["company"]["code"]
    interim = path(cfg, "interim_dir")
    output = path(cfg, "output_dir")
    cols = cfg["columns"]

    top_only_env = os.environ.get("EXPORT_TOP_ONLY", "").strip()
    top_only = int(top_only_env) if top_only_env.isdigit() else None
    use_csv = os.environ.get("EXPORT_CSV", "").strip() in {"1", "true", "yes"}

    src = interim / f"{company}_tagged_rows.parquet"
    if not src.exists():
        raise FileNotFoundError(f"Run step 4 first. Missing: {src}")

    print(f"Loading {src} ...", flush=True)
    df = pd.read_parquet(src)
    n = len(df)
    print(f"  rows={n:,}", flush=True)

    keep_cols_priority = [
        cols.get("period", ""),
        cols.get("posting_year", ""),
        cols.get("posting_date", ""),
        cols.get("ng11_category", ""),
        cols.get("project", ""),
        cols.get("capex_opex", ""),
        cols.get("amount", ""),
        cols.get("account_code", ""),
        cols.get("account_desc", ""),
        cols.get("description", ""),
        cols.get("vendor", ""),
        "signature",
        "vertical_id",
        "vertical_label",
        "vertical_source",
        "horizontal_id",
        "horizontal_label",
        "horizontal_source",
        "final_capex_opex",
        "ng_scope",
        "row_type",
    ]
    keep_cols = [c for c in keep_cols_priority if c and c in df.columns]
    df = df[keep_cols].copy()

    if cols.get("amount") and cols["amount"] in df.columns:
        df["_abs"] = pd.to_numeric(df[cols["amount"]], errors="coerce").abs()
        df = df.sort_values("_abs", ascending=False, na_position="last")
        df = df.drop(columns=["_abs"])

    if top_only:
        print(f"  filtering to top {top_only} by abs(amount)", flush=True)
        df = df.head(top_only)

    suffix = "_tagged_rows"
    if top_only:
        suffix += f"_top{top_only}"

    if use_csv:
        out_path = output / f"{company}{suffix}.csv"
        print(f"Writing CSV: {out_path} ...", flush=True)
        df.to_csv(out_path, index=False, encoding="utf-8-sig")
    else:
        out_path = output / f"{company}{suffix}.xlsx"
        print(f"Writing xlsx (this can take 30-60s for 600k rows): {out_path} ...", flush=True)
        with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name="tagged_rows", index=False)
    size_mb = out_path.stat().st_size / 1e6
    print(f"\nDone. {len(df):,} rows  {size_mb:.1f} MB  ->  {out_path}")
