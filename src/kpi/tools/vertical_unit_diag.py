"""Diagnose whether `project` alone is unique enough as the 縱向 classification unit.

Some projects span multiple NG categories or mix CAPEX/OPEX in ways that should NOT
share a single vertical tag. This tool counts:

  - unique projects
  - unique (project, ng11_category) combos
  - projects that have >1 NG value  → candidates for splitting
  - projects that have >1 capex_opex value
  - projects with mixed-sign amounts (positive + negative)

Usage (from repo root, after step0 has produced parquets):
  python -m kpi.tools.vertical_unit_diag                  # all configured entities
  python -m kpi.tools.vertical_unit_diag company_3        # one entity

Output: stdout summary + optional Excel detail file at
  data/audit/vertical_unit_diag_<entity>.xlsx
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import yaml

from kpi.lib.conf import ROOT
from kpi.lib.io_setup import force_unbuffered_io

force_unbuffered_io()


CONF = ROOT / "conf"
AUDIT_DIR = ROOT / "data" / "audit"
TOP_N = 20


def _load_entity(entity_key: str) -> dict | None:
    f = CONF / entity_key / "parameters.yml"
    if not f.exists():
        return None
    with f.open(encoding="utf-8") as fh:
        return yaml.safe_load(fh)


def _all_entities() -> list[str]:
    return sorted(p.name for p in CONF.iterdir()
                  if p.is_dir() and p.name.startswith("company_"))


def diag_one(entity_key: str, write_xlsx: bool = True) -> None:
    cfg = _load_entity(entity_key)
    if not cfg:
        print(f"\n[{entity_key}] no parameters.yml — skip")
        return
    alias = cfg.get("alias", entity_key)
    cols = cfg.get("columns") or {}
    parquet = ROOT / "data" / alias / "interim" / f"{entity_key}_raw.parquet"

    proj_col = cols.get("project")
    ng_col = cols.get("ng11_category")
    co_col = cols.get("capex_opex")
    amt_col = cols.get("amount")

    print(f"\n══ {entity_key} ({alias}) ══")
    if not parquet.exists():
        print(f"  parquet missing: {parquet} — run step0 first")
        return
    if not (proj_col and ng_col and amt_col):
        print(f"  required columns not configured (project/ng11_category/amount)")
        return

    df = pd.read_parquet(parquet, columns=[c for c in [proj_col, ng_col, co_col, amt_col] if c])
    df = df.rename(columns={proj_col: "_proj", ng_col: "_ng", amt_col: "_amt"})
    if co_col and co_col in df.columns:
        df = df.rename(columns={co_col: "_co"})
    else:
        df["_co"] = ""
    df["_amt"] = pd.to_numeric(df["_amt"], errors="coerce").fillna(0)
    df["_proj"] = df["_proj"].astype("string").fillna("(空白)")
    df["_ng"] = df["_ng"].astype("string").fillna("(空白)")
    df["_co"] = df["_co"].astype("string").fillna("(空白)")

    n_rows = len(df)
    n_proj = df["_proj"].nunique()
    n_proj_ng = df.groupby(["_proj", "_ng"], dropna=False).ngroups
    n_proj_co = df.groupby(["_proj", "_co"], dropna=False).ngroups

    g = df.groupby("_proj", dropna=False)
    ng_per_proj = g["_ng"].nunique()
    co_per_proj = g["_co"].nunique()
    pos_per_proj = g.apply(lambda x: (x["_amt"] > 0).any())
    neg_per_proj = g.apply(lambda x: (x["_amt"] < 0).any())

    multi_ng = ng_per_proj[ng_per_proj > 1]
    multi_co = co_per_proj[co_per_proj > 1]
    mixed_sign = pos_per_proj & neg_per_proj
    mixed_sign_n = int(mixed_sign.sum())

    print(f"  rows: {n_rows:,}")
    print(f"  unique projects:           {n_proj:,}")
    print(f"  unique (project, ng):      {n_proj_ng:,}   ← granularity if NG split added")
    print(f"  unique (project, capex):   {n_proj_co:,}")
    print(f"  projects w/ >1 NG:         {len(multi_ng):,}  ({len(multi_ng)/max(n_proj,1)*100:.1f}%)")
    print(f"  projects w/ >1 capex_opex: {len(multi_co):,}  ({len(multi_co)/max(n_proj,1)*100:.1f}%)")
    print(f"  projects w/ mixed-sign $:  {mixed_sign_n:,}  (zero-net candidates)")

    if len(multi_ng) > 0:
        print(f"\n  Top {min(TOP_N, len(multi_ng))} projects with multiple NG values (by |amount|):")
        amt_by_proj = g["_amt"].apply(lambda s: float(s.abs().sum()))
        ranked = (
            multi_ng.to_frame("ng_count")
            .join(amt_by_proj.rename("abs_amt"))
            .sort_values("abs_amt", ascending=False)
            .head(TOP_N)
        )
        for proj, row in ranked.iterrows():
            ngs = sorted(df[df["_proj"] == proj]["_ng"].dropna().unique().tolist())
            print(f"    [{int(row['ng_count'])} NGs, |${row['abs_amt']:,.0f}|]  {proj[:55]:<55}  → {ngs}")

    if write_xlsx and len(multi_ng) > 0:
        AUDIT_DIR.mkdir(parents=True, exist_ok=True)
        out = AUDIT_DIR / f"vertical_unit_diag_{entity_key}.xlsx"
        with pd.ExcelWriter(out, engine="xlsxwriter") as w:
            ranked_full = (
                multi_ng.to_frame("ng_count")
                .join(g["_amt"].apply(lambda s: float(s.abs().sum())).rename("abs_amt"))
                .sort_values("abs_amt", ascending=False)
            )
            ranked_full.reset_index().to_excel(w, sheet_name="multi_ng_projects", index=False)
            pn = (
                df.groupby(["_proj", "_ng"], dropna=False)
                .agg(rows=("_amt", "size"), amount=("_amt", "sum"))
                .reset_index()
            )
            pn.to_excel(w, sheet_name="project_x_ng", index=False)
        print(f"  wrote: {out.relative_to(ROOT)}")


def main() -> None:
    args = sys.argv[1:]
    targets = args if args else _all_entities()
    for ent in targets:
        diag_one(ent, write_xlsx=True)


if __name__ == "__main__":
    main()
