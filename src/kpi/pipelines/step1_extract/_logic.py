"""Step 1: Extract unique projects, accounts, vendors, and row signatures for tagging review.

A "signature" = (account_code, account_desc, normalized(description)) — used as the
unit for 橫向 tagging because account_code alone misses context (refunds, adjustments,
context-dependent spend). Signatures collapse 624k rows down to thousands of LLM-able
units while preserving description-level nuance.

Outputs:
  data/interim/<company>_unique_projects.xlsx
  data/interim/<company>_unique_accounts.xlsx
  data/interim/<company>_unique_vendors.xlsx
  data/interim/<company>_unique_signatures.xlsx
  data/interim/<company>_summary.txt
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

from kpi.lib.conf import load_config, path  # noqa: E402
from kpi.lib.io_setup import force_unbuffered_io  # noqa: E402
from kpi.lib.text import normalize_description  # noqa: E402

force_unbuffered_io()


def main():
    cfg = load_config()
    interim = path(cfg, "interim_dir")
    company = cfg["company"]["code"]
    src = interim / f"{company}_raw.parquet"
    if not src.exists():
        raise FileNotFoundError(f"Run step0 first. Missing: {src}")

    sig_path = interim / f"{company}_unique_signatures.xlsx"
    proj_path = interim / f"{company}_unique_projects.xlsx"
    acct_path = interim / f"{company}_unique_accounts.xlsx"
    vend_path = interim / f"{company}_unique_vendors.xlsx"
    # Gate each output INDEPENDENTLY. The old behaviour skipped the whole node if the sig file
    # existed, so rebuilding unique_projects (e.g. to pick up a newly-added year's projects) forced
    # deleting unique_signatures → a full step3 re-LLM. Now: rebuild whichever project/account/vendor
    # file is missing, and MERGE new signatures into the existing sig file (existing rows + their
    # step3 tags preserved; only genuinely new sigs appended). Skip the node only if ALL exist.
    _need_vend = bool((cfg["columns"] or {}).get("vendor"))
    if (proj_path.exists() and acct_path.exists() and sig_path.exists()
            and (vend_path.exists() or not _need_vend)):
        print("  [skip] all step1 outputs exist — delete a specific *_unique_*.xlsx to rebuild it "
              "(new signatures merge automatically; unique_signatures is never wiped)")
        return

    cols = cfg["columns"]
    df = pd.read_parquet(src)
    n = len(df)

    def agg(group_col: str, extra_cols: list[str]) -> pd.DataFrame:
        gb = df.groupby(group_col, dropna=False)
        out = gb[cols["amount"]].agg(["count", "sum"]).rename(
            columns={"count": "row_count", "sum": "total_amount"}
        )
        for c in extra_cols:
            if c in df.columns:
                out[c + "__sample"] = gb[c].agg(
                    lambda s: " | ".join(map(str, pd.Series(s.dropna().unique())[:3]))
                )
        return out.reset_index().sort_values("total_amount", ascending=False, na_position="last")

    proj = agg(cols["project"], [cols["ng11_category"], cols["capex_opex"], cols["account_desc"], cols["description"]])
    proj["manual_vertical"] = ""
    proj["manual_capex_opex"] = ""
    proj["llm_vertical"] = ""
    proj["llm_capex_opex"] = ""
    proj["llm_confidence"] = ""
    proj["llm_reasoning"] = ""

    acct = agg(cols["account_code"], [cols["account_desc"]])
    acct["manual_horizontal"] = ""
    acct["rule_horizontal"] = ""
    acct["llm_horizontal"] = ""

    vend = agg(cols["vendor"], [cols["account_desc"], cols["description"]]) if cols.get("vendor") else pd.DataFrame()

    sig_df = build_signatures(df, cols)

    # Write whichever project/account/vendor file is missing — don't clobber an existing tagged file.
    if not proj_path.exists():
        proj.to_excel(proj_path, index=False)
    if not acct_path.exists():
        acct.to_excel(acct_path, index=False)
    if cols.get("vendor") and not vend_path.exists():
        vend.to_excel(vend_path, index=False)
    # Signatures: create if missing, else MERGE — keep every existing row (and its step3 columns /
    # tags) and only APPEND signatures not already present. So adding a year never wipes the sig file
    # nor forces a full step3 re-LLM; only the genuinely new sigs are left untagged for step3/feedback.
    if not sig_path.exists():
        sig_df.to_excel(sig_path, index=False)
        print(f"  [sig] created {sig_path.name} with {len(sig_df):,} signatures")
    else:
        _existing = pd.read_excel(sig_path)
        _ekeys = set(_existing["signature"].astype(str)) if "signature" in _existing.columns else set()
        _new = sig_df[~sig_df["signature"].astype(str).isin(_ekeys)].copy()
        if len(_new):
            _new = _new.reindex(columns=_existing.columns)
            pd.concat([_existing, _new], ignore_index=True).to_excel(sig_path, index=False)
            print(f"  [sig] merged {len(_new):,} NEW signatures into {sig_path.name} "
                  f"(kept {len(_existing):,} existing + their tags)")
        else:
            print(f"  [sig] no new signatures ({len(_existing):,} existing kept)")

    summary_lines = [
        f"company: {company}",
        f"total rows: {n:,}",
        f"unique projects: {df[cols['project']].nunique(dropna=False):,}",
        f"unique account codes: {df[cols['account_code']].nunique(dropna=False):,}",
        f"unique vendors: {df[cols['vendor']].nunique(dropna=False):,}" if cols.get("vendor") else "unique vendors: (not mapped)",
        f"unique row signatures: {len(sig_df):,}",
        f"signatures with negative amount: {int((sig_df['has_negative']).sum()):,}",
        f"period values: {sorted(map(str, df[cols['period']].dropna().unique().tolist())) if cols.get('period') and cols['period'] in df.columns else '(not mapped)'}",
        f"posting_year range: {df[cols['posting_year']].min()} .. {df[cols['posting_year']].max()}" if cols.get('posting_year') and cols['posting_year'] in df.columns else "posting_year range: (not mapped)",
        f"ng11 category values: {sorted(map(str, df[cols['ng11_category']].dropna().unique().tolist()))}",
        f"capex/opex values: {sorted(map(str, df[cols['capex_opex']].dropna().unique().tolist()))}",
        f"total amount: {df[cols['amount']].sum():,.0f}",
        f"rows with amount==null: {df[cols['amount']].isna().sum():,}",
        f"rows with negative amount: {int((df[cols['amount']] < 0).sum()):,}",
    ]
    summary = "\n".join(summary_lines)
    (interim / f"{company}_summary.txt").write_text(summary, encoding="utf-8")
    print(summary)
    print(f"\nWrote:\n  {proj_path}\n  {acct_path}\n  {vend_path}\n  {sig_path}")


_RESERVED_COL_KEYS = {
    "period", "posting_year", "posting_date", "amount", "capex_opex",
    "project", "project_name_cols", "ng11_category",
    "account_code", "account_code_cols",
    "account_desc",
    "description", "description_cols",
    "vendor", "job_code", "unique_id", "unique_id_cols", "count_cols",
}


def build_signatures(df: pd.DataFrame, cols: dict) -> pd.DataFrame:
    desc_col = cols["description"]
    acct_col = cols["account_code"]
    adesc_col = cols["account_desc"]
    amt_col = cols["amount"]
    proj_col = cols["project"]
    co_col = cols.get("capex_opex", "")
    jc_col = cols.get("job_code", "")
    jc_configured = bool(jc_col) and jc_col in df.columns

    work = pd.DataFrame({
        "account_code": df[acct_col].astype("string").fillna("").str.strip(),
        "account_desc": df[adesc_col].astype("string").fillna("").str.strip(),
        "desc_norm": df[desc_col].apply(normalize_description),
        "desc_raw": df[desc_col].astype("string").fillna(""),
        "amount": pd.to_numeric(df[amt_col], errors="coerce"),
        "project": df[proj_col].astype("string").fillna(""),
        "capex_opex": (
            df[co_col].astype("string").fillna("").str.strip()
            if co_col and co_col in df.columns else pd.Series([""] * len(df))
        ),
        "job_code": (
            df[jc_col].astype("string").fillna("").str.strip()
            if jc_configured else pd.Series([""] * len(df))
        ),
    })

    # Auto-detect "extra" columns: any string-valued column key in `columns:` that's
    # NOT in RESERVED is treated as a user-defined extra field, surfaced as
    # work["<key>"] for aggregation and rule-predicate lookup.
    extra_keys: list[str] = []
    for k, v in cols.items():
        if k in _RESERVED_COL_KEYS:
            continue
        if not isinstance(v, str) or not v:
            continue
        if v in df.columns:
            work[k] = df[v].astype("string").fillna("").str.strip()
            extra_keys.append(k)

    # Signature: include job_code only when configured, so non-affected companies'
    # signature text (and hence step3 LLM cache keys) stay unchanged.
    if jc_configured:
        work["signature"] = (
            work["account_code"] + "|" + work["account_desc"] + "|" + work["desc_norm"] + "|" + work["job_code"]
        )
    else:
        work["signature"] = work["account_code"] + "|" + work["account_desc"] + "|" + work["desc_norm"]

    grp = work.groupby("signature", dropna=False)
    out = pd.DataFrame({
        "signature": grp.size().index,
        "row_count": grp.size().values,
        "total_amount": grp["amount"].sum().values,
        "amount_min": grp["amount"].min().values,
        "amount_max": grp["amount"].max().values,
        "neg_row_count": grp.apply(lambda g: int((g["amount"] < 0).sum())).values,
    })
    out["has_negative"] = out["neg_row_count"] > 0

    # Dominant Capex/Opex per signature (helps step3 LLM bias H_CONSTRUCTION when Capex)
    def _co_dominant(s: pd.Series) -> str:
        vc = s.replace("", pd.NA).dropna().value_counts()
        return str(vc.index[0]) if len(vc) > 0 else ""
    def _co_pct(s: pd.Series) -> int:
        vc = s.replace("", pd.NA).dropna().value_counts(normalize=True)
        return int(round(vc.iloc[0] * 100)) if len(vc) > 0 else 0
    out["capex_opex_dominant"] = grp["capex_opex"].agg(_co_dominant).values
    out["capex_opex_dominant_pct"] = grp["capex_opex"].agg(_co_pct).values

    first_meta = grp.first()[["account_code", "account_desc", "desc_norm"]].reset_index(drop=True)
    out = out.reset_index(drop=True).join(first_meta)

    desc_samples = grp["desc_raw"].agg(
        lambda s: " || ".join(pd.Series([x for x in s if x]).drop_duplicates().head(3).tolist())
    ).reset_index(drop=True)
    out["desc_samples"] = desc_samples

    proj_samples = grp["project"].agg(
        lambda s: " || ".join(pd.Series([x for x in s if x]).drop_duplicates().head(3).tolist())
    ).reset_index(drop=True)
    out["project_samples"] = proj_samples
    out["project_distinct_count"] = grp["project"].nunique().reset_index(drop=True)

    # Job code aggregation — all unique values per signature (no head() limit so the
    # predicate `job_code_contains` won't miss values truncated by sampling).
    jc_all = grp["job_code"].agg(
        lambda s: " || ".join(pd.Series([x for x in s if x]).drop_duplicates().tolist())
    ).reset_index(drop=True)
    out["job_code_samples"] = jc_all

    # Aggregate every auto-detected extra column as `<key>_samples`.
    for k in extra_keys:
        out[f"{k}_samples"] = grp[k].agg(
            lambda s: " || ".join(pd.Series([x for x in s if x]).drop_duplicates().head(20).tolist())
        ).reset_index(drop=True)

    base_cols = [
        "signature",
        "account_code",
        "account_desc",
        "desc_norm",
        "desc_samples",
        "row_count",
        "total_amount",
        "amount_min",
        "amount_max",
        "neg_row_count",
        "has_negative",
        "capex_opex_dominant",
        "capex_opex_dominant_pct",
        "project_distinct_count",
        "project_samples",
        "job_code_samples",
    ]
    out = out[base_cols + [f"{k}_samples" for k in extra_keys]].sort_values(
        "total_amount", ascending=False, na_position="last"
    )

    out["manual_horizontal"] = ""
    out["manual_row_type"] = ""
    out["llm_horizontal"] = ""
    out["llm_row_type"] = ""
    out["llm_confidence"] = ""
    out["llm_reasoning"] = ""
    return out


if __name__ == "__main__":
    main()
