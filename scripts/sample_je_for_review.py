"""JE-level (transaction) sampler for human review.

For a given (entity, year), sample N JE rows from each (NG × capex × LLM_H)
stratification cell. Output a single table-like txt for paste-back review.

Why JE-level review matters:
  - Sig-level summary hides cell-specific mis-classifications
  - Real ground truth lives in individual transaction context (acct/desc/amt)
  - Helps spot systematic LLM errors (e.g. "Comp Leave" → wrong H)

Output: je_review_{ent}_{year}.txt

Columns shown per row:
  NG  | capex/opex  | LLM_H  | amount  | acct_desc  | description  | project  | company

Stratification strategy:
  - For each NG (NG0..NG11), sample top-N high-|amount| rows
  - Within each NG, balance capex vs opex if both exist
  - Optionally filter to rows currently tagged with a specific H (e.g. H_OTHER for refinement)

Run:
  python scripts/sample_je_for_review.py --entity galaxy --year 25
  python scripts/sample_je_for_review.py --entity galaxy --year 25 --per-ng 8
  python scripts/sample_je_for_review.py --entity galaxy --year 25 --filter-h H_OTHER
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import pandas as pd
import yaml


ENTITIES = {
    "galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
    "vml":    "company_4", "melco":  "company_5", "mgm": "company_6",
}


def _find_year_col(df: pd.DataFrame) -> str | None:
    for c in ("report_period", "report_year", "Yr related", "years"):
        if c in df.columns:
            return c
    return None


def _filter_year(df: pd.DataFrame, year_col: str, year: str) -> pd.DataFrame:
    s = df[year_col].astype(str)
    mask = (
        s.str.startswith(year)
        | (s == f"Yr 20{year}")
        | (s == f"20{year}")
        | (s == f"20{year}-01-01 00:00:00")
        | s.str.contains(f"20{year}年", na=False)
    )
    return df[mask].copy()


def _find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    for c in candidates:
        if c in df.columns:
            return c
    return None


def sample_entity(ent: str, com: str, year: str, per_ng: int = 8,
                  filter_h: str | None = None) -> list[str]:
    out: list[str] = []
    report_parquet = Path(f"data/{ent}/output/{com}_kpi_report.parquet")
    raw_parquet = Path(f"data/{ent}/interim/{com}_raw.parquet")

    if not report_parquet.exists():
        return [f"[{ent}-{year}] no kpi_report.parquet — kedro step5 not done"]

    cfg = yaml.safe_load(Path(f"conf/{com}/parameters.yml").read_text(encoding="utf-8"))
    cols_cfg = cfg.get("columns", {})
    amt_col = cols_cfg.get("amount", "")
    ac_col = cols_cfg.get("account_code", "")
    ad_col = cols_cfg.get("account_desc", "")
    dn_col = cols_cfg.get("description", "")
    proj_col = cols_cfg.get("project", "")

    df = pd.read_parquet(report_parquet)
    year_col = _find_year_col(df)
    if not year_col:
        return [f"[{ent}-{year}] no year col in report parquet (cols: {list(df.columns)[:8]}...)"]

    df = _filter_year(df, year_col, year)
    if len(df) == 0:
        return [f"[{ent}-{year}] no rows after year filter"]

    # Detect NG col + capex/opex col
    ng_col = _find_col(df, ["NG11 Category", "NG11 category", "ng11_category", "ng_scope"])
    capex_col = _find_col(df, ["final_capex_opex", "Capex/Opex", "CAPEX/OPEX",
                                "Capex / Opex", "Ledger Type", "Capex/Opex重分類", "item_type"])
    h_col = "horizontal_id" if "horizontal_id" in df.columns else None
    company_col = _find_col(df, ["Company Name", "company", "公司簡稱", "投資主體名稱"])
    vendor_col = _find_col(df, ["Vendor Name", "Vendor", "vendor"])

    if not (h_col and amt_col in df.columns):
        return [f"[{ent}-{year}] missing horizontal_id or amount col"]

    df["_amt"] = pd.to_numeric(df[amt_col], errors="coerce").fillna(0)
    df["_abs"] = df["_amt"].abs()

    out.append("=" * 130)
    out.append(f"JE SAMPLE — {ent}-{year}  total_rows={len(df):,}  total_amount={df['_amt'].sum():,.0f}")
    out.append(f"  per_ng={per_ng}  filter_h={filter_h or '(none)'}")
    out.append(f"  cols: ng='{ng_col}'  capex='{capex_col}'  h='{h_col}'  acct_desc='{ad_col}'  desc='{dn_col}'")
    out.append("=" * 130)

    if filter_h:
        df = df[df[h_col].astype(str) == filter_h]
        out.append(f"  After H filter: {len(df):,} rows")

    if not ng_col:
        out.append(f"  ⚠️  no NG col found — sampling top {per_ng*12} rows by amount instead")
        top = df.sort_values("_abs", ascending=False).head(per_ng * 12)
        ng_groups = [("(no NG col)", top)]
    else:
        ng_values = df[ng_col].astype(str).unique()
        # Order: NG0, NG1, ..., NG11
        ng_order = sorted(ng_values, key=lambda x: (0, int(x[2:])) if x.startswith("NG") and x[2:].isdigit() else (1, x))
        ng_groups = []
        for ng in ng_order:
            sub = df[df[ng_col].astype(str) == ng]
            if len(sub) == 0:
                continue
            # Balance capex / opex within NG
            if capex_col and capex_col in sub.columns:
                ce_vals = sub[capex_col].astype(str).str.lower()
                capex_rows = sub[ce_vals.str.contains("capex", na=False)]
                opex_rows = sub[ce_vals.str.contains("opex", na=False)]
                per_side = max(1, per_ng // 2)
                capex_sample = capex_rows.sort_values("_abs", ascending=False).head(per_side)
                opex_sample = opex_rows.sort_values("_abs", ascending=False).head(per_side)
                ng_sample = pd.concat([capex_sample, opex_sample]).head(per_ng)
            else:
                ng_sample = sub.sort_values("_abs", ascending=False).head(per_ng)
            ng_groups.append((ng, ng_sample))

    # Render
    for ng, sub in ng_groups:
        if len(sub) == 0:
            continue
        ng_total_amt = float(sub["_amt"].sum())
        out.append(f"\n## {ng}  ({len(sub)} sampled  amount={ng_total_amt:,.0f}) ##")
        out.append(f"  {'idx':>4} {'capex':<6} {'LLM_H':<18} {'amount':>15}  acct_desc | description | project | company")
        out.append("  " + "-" * 130)
        for idx, (_, r) in enumerate(sub.iterrows(), 1):
            ce_val = str(r[capex_col])[:5] if capex_col and capex_col in sub.columns else "?"
            llm_h = str(r[h_col])[:18]
            ad_v = str(r.get(ad_col, ""))[:35] if ad_col else ""
            dn_v = str(r.get(dn_col, ""))[:55] if dn_col else ""
            pj_v = str(r.get(proj_col, ""))[:40] if proj_col else ""
            co_v = str(r.get(company_col, ""))[:25] if company_col else ""
            out.append(f"  {idx:>4} {ce_val:<6} {llm_h:<18} {r['_amt']:>15,.0f}  {ad_v} | {dn_v} | {pj_v} | {co_v}")

    return out


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--entity", choices=list(ENTITIES), required=True)
    parser.add_argument("--year", default="25")
    parser.add_argument("--per-ng", type=int, default=8, help="rows to sample per NG (default 8)")
    parser.add_argument("--filter-h", help="only sample rows tagged this H (e.g. H_OTHER)")
    args = parser.parse_args()

    com = ENTITIES[args.entity]
    print(f">>> sampling {args.entity}-{args.year}...", flush=True)
    lines = sample_entity(args.entity, com, args.year, args.per_ng, args.filter_h)

    suffix = f"_{args.filter_h}" if args.filter_h else ""
    out_path = Path(f"je_review_{args.entity}_{args.year}{suffix}.txt")
    out_path.write_text("\n".join(lines), encoding="utf-8")
    for line in lines:
        print(line, flush=True)
    print(f"\n✓ wrote {out_path}")


if __name__ == "__main__":
    main()
