"""Cross-company audit sampler — single output covering all 6 entities.

Goal: a small, hand-reviewable file Claude / project team can use to
identify prompt-tuning opportunities. Pulls representative slices from each
entity's tagged_rows.parquet + unique_signatures.xlsx:

  rows_h_other       — H_OTHER tagged rows (top |amount| per company)
  rows_v_other       — V_OTHER tagged rows
  rows_top_amount    — biggest |amount| rows per company
  rows_random        — random cross-section
  sigs_low_conf      — signatures with LLM confidence < 0.7
  sigs_h_other       — signatures tagged H_OTHER
  projects_v_other   — unique projects tagged V_OTHER
  projects_low_conf  — projects with LLM confidence < 0.7

Output: data/audit/audit_review_<YYYYMMDD>.xlsx

NOT per-entity — runs once and reads all 6 (or however many have output).
"""
from __future__ import annotations

import datetime as dt
from pathlib import Path

import pandas as pd

from kpi.lib.conf import ROOT, folder_for, load_master
from kpi.lib.io_setup import force_unbuffered_io

force_unbuffered_io()


AUDIT_DIR = ROOT / "data" / "audit"
SAMPLE_PER_COMPANY = 30
LOW_CONF = 0.7

ROW_COLS_CANDIDATES = [
    "_company", "period", "posting_year",
    "project", "ng11_category", "capex_opex", "amount",
    "account_code", "account_desc", "description", "vendor",
    "vertical_label", "vertical_source",
    "horizontal_label", "horizontal_source",
    "final_capex_opex", "row_type", "ng_scope",
]

SIG_COLS_CANDIDATES = [
    "company", "account_code", "account_desc", "desc_norm", "desc_samples",
    "row_count", "total_amount", "capex_opex_dominant",
    "llm_horizontal", "llm_row_type", "llm_confidence", "llm_reasoning",
    "manual_horizontal",
]

PROJ_COLS_CANDIDATES = [
    "company", "primary_ng", "row_count", "total_amount",
    "llm_vertical", "llm_vertical_label", "llm_confidence",
    "llm_reasoning", "llm_alt_candidates",
    "tag_source", "manual_vertical",
]


def _load_rows(entity_key: str, alias: str, cols_map: dict) -> pd.DataFrame:
    parquet = ROOT / "data" / alias / "interim" / f"{entity_key}_tagged_rows.parquet"
    if not parquet.exists():
        return pd.DataFrame()
    df = pd.read_parquet(parquet)
    df["_company"] = alias
    amt_col = cols_map.get("amount")
    if amt_col and amt_col in df.columns:
        df["_amount"] = pd.to_numeric(df[amt_col], errors="coerce").fillna(0)
        df["_abs"] = df["_amount"].abs()
        df["amount"] = df["_amount"]
    else:
        df["_abs"] = 0.0
    # Project the canonical column names so different entities' columns line up.
    for canon in ("period", "posting_year", "project", "ng11_category",
                  "capex_opex", "account_code", "account_desc", "description", "vendor"):
        src_col = cols_map.get(canon)
        if src_col and src_col in df.columns and canon not in df.columns:
            df[canon] = df[src_col]
    return df


def _load_signatures(entity_key: str, alias: str) -> pd.DataFrame:
    p = ROOT / "data" / alias / "interim" / f"{entity_key}_unique_signatures.xlsx"
    if not p.exists():
        return pd.DataFrame()
    df = pd.read_excel(p)
    df["company"] = alias
    df["abs_amount"] = pd.to_numeric(df.get("total_amount"), errors="coerce").abs()
    df["llm_confidence"] = pd.to_numeric(df.get("llm_confidence"), errors="coerce")
    return df


def _load_projects(entity_key: str, alias: str) -> pd.DataFrame:
    p = ROOT / "data" / alias / "interim" / f"{entity_key}_unique_projects.xlsx"
    if not p.exists():
        return pd.DataFrame()
    df = pd.read_excel(p)
    df["company"] = alias
    df["abs_amount"] = pd.to_numeric(df.get("total_amount"), errors="coerce").abs()
    df["llm_confidence"] = pd.to_numeric(df.get("llm_confidence"), errors="coerce")
    return df


def _write_sample(writer, sheet: str, df: pd.DataFrame, candidate_cols: list[str]) -> int:
    if df.empty:
        return 0
    cols = [c for c in candidate_cols if c in df.columns]
    df[cols].to_excel(writer, sheet_name=sheet[:31], index=False)
    return len(df)


def main() -> None:
    master = load_master()
    companies = master.get("companies") or {}

    # Buckets: list of DataFrames per category, concatenated at the end.
    row_buckets: dict[str, list[pd.DataFrame]] = {
        "h_other": [], "v_other": [], "top_amount": [], "random": [],
    }
    sig_buckets: dict[str, list[pd.DataFrame]] = {"low_conf": [], "h_other": []}
    proj_buckets: dict[str, list[pd.DataFrame]] = {"v_other": [], "low_conf": []}

    found = 0
    for entity_key, ent in companies.items():
        if not ent or not ent.get("name"):
            continue
        alias = folder_for(entity_key, ent)
        cols_map = ent.get("columns") or {}

        rows = _load_rows(entity_key, alias, cols_map)
        sigs = _load_signatures(entity_key, alias)
        projs = _load_projects(entity_key, alias)
        if rows.empty and sigs.empty and projs.empty:
            print(f"  [skip] {entity_key} ({alias}) — no audit data yet")
            continue
        found += 1
        print(f"  -- {entity_key} ({alias}): rows={len(rows):,}  sigs={len(sigs)}  projs={len(projs)}")

        if not rows.empty:
            h_other = rows[rows.get("horizontal_label", pd.Series(dtype="object")) == "H_OTHER"]
            v_other = rows[rows.get("vertical_label", pd.Series(dtype="object")) == "V_OTHER"]
            top_amt = rows.sort_values("_abs", ascending=False)
            rnd = rows.sample(min(SAMPLE_PER_COMPANY, len(rows)), random_state=42)
            row_buckets["h_other"].append(h_other.sort_values("_abs", ascending=False).head(SAMPLE_PER_COMPANY))
            row_buckets["v_other"].append(v_other.sort_values("_abs", ascending=False).head(SAMPLE_PER_COMPANY))
            row_buckets["top_amount"].append(top_amt.head(SAMPLE_PER_COMPANY))
            row_buckets["random"].append(rnd)

        if not sigs.empty:
            low_conf = sigs[sigs["llm_confidence"].fillna(1.0) < LOW_CONF]
            low_conf = low_conf.sort_values("abs_amount", ascending=False).head(SAMPLE_PER_COMPANY)
            sig_buckets["low_conf"].append(low_conf)
            h_other_sigs = sigs[sigs.get("llm_horizontal", "") == "H_OTHER"]
            h_other_sigs = h_other_sigs.sort_values("abs_amount", ascending=False).head(SAMPLE_PER_COMPANY)
            sig_buckets["h_other"].append(h_other_sigs)

        if not projs.empty:
            v_other = projs[projs.get("llm_vertical", "") == "V_OTHER"]
            v_other = v_other.sort_values("abs_amount", ascending=False).head(SAMPLE_PER_COMPANY)
            proj_buckets["v_other"].append(v_other)
            low_conf = projs[projs["llm_confidence"].fillna(1.0) < LOW_CONF]
            low_conf = low_conf.sort_values("abs_amount", ascending=False).head(SAMPLE_PER_COMPANY)
            proj_buckets["low_conf"].append(low_conf)

    if found == 0:
        print("No company audit data found anywhere — run main pipelines first.")
        return

    AUDIT_DIR.mkdir(parents=True, exist_ok=True)
    today = dt.date.today().strftime("%Y%m%d")
    out_path = AUDIT_DIR / f"audit_review_{today}.xlsx"

    summary_rows = []
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        for bucket, dfs in row_buckets.items():
            n = _write_sample(
                writer,
                f"rows_{bucket}",
                pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame(),
                ROW_COLS_CANDIDATES,
            )
            summary_rows.append({"sheet": f"rows_{bucket}", "rows": n})
        for bucket, dfs in sig_buckets.items():
            n = _write_sample(
                writer,
                f"sigs_{bucket}",
                pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame(),
                SIG_COLS_CANDIDATES,
            )
            summary_rows.append({"sheet": f"sigs_{bucket}", "rows": n})
        for bucket, dfs in proj_buckets.items():
            n = _write_sample(
                writer,
                f"projects_{bucket}",
                pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame(),
                PROJ_COLS_CANDIDATES,
            )
            summary_rows.append({"sheet": f"projects_{bucket}", "rows": n})

    print(f"\nWrote {out_path.relative_to(ROOT)}")
    for r in summary_rows:
        print(f"  {r['sheet']:25s}  {r['rows']:>5} rows")
    print(f"\n  companies covered: {found}")
