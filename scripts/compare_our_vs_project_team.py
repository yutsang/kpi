"""Compare our (LLM/rules) classification vs project team manual labels.

Goal:
  1. Find where we disagree with project team (ground truth) — fix THIS run.
  2. Extract patterns to teach LLM — improve NEXT run's accuracy.

For each row that has a project-team manual label:
  - Map the manual label → expected_H (via MAPPINGS)
  - Compare with our tagged horizontal_id (from kpi_report.parquet)
  - Group disagreements by (expected_H, our_H)

Output per entity:
  data/{ent}_label_disagreement.txt
    SECTION A: overall agreement rate (rows + amount)
    SECTION B: per (expected_H, our_H) disagreement bucket
               + top sigs in that bucket (so we know what's getting mislabeled)
               + suggested system-prompt addition
    SECTION C: rows with manual label that we tagged H_OTHER
               (high-value fixes — applying manual label gives instant improvement)
    SECTION D: LLM-teaching patterns
               (account/desc text patterns that should map to expected_H)

Run:
  python scripts/compare_our_vs_project_team.py --all
  python scripts/compare_our_vs_project_team.py --entity wynn
"""
from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import pandas as pd
import yaml

from apply_manual_labels_as_overrides import MAPPINGS
from kpi.lib.text import normalize_description


ENTITIES = {
    "galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
    "vml":    "company_4", "melco":  "company_5", "mgm": "company_6",
}

EMPTY_VALUES = {"", "nan", "<na>", "-", "0", "none", "na", "null", "n/a"}


def is_emptyish_series(s: pd.Series) -> pd.Series:
    return s.isna() | s.astype(str).str.strip().str.lower().isin(EMPTY_VALUES)


def _build_sig(df: pd.DataFrame, cfg) -> pd.Series:
    cols = cfg.get("columns", {})
    ac = cols.get("account_code", "")
    ad = cols.get("account_desc", "")
    dn = cols.get("description", "")
    jc = cols.get("job_code", "")
    a = df[ac].astype("string").fillna("").str.strip() if ac in df.columns else ""
    b = df[ad].astype("string").fillna("").str.strip() if ad in df.columns else ""
    c = df[dn].apply(normalize_description) if dn in df.columns else ""
    if jc and jc in df.columns:
        d = df[jc].astype("string").fillna("").str.strip()
        return a + "|" + b + "|" + c + "|" + d
    return a + "|" + b + "|" + c


def compare_entity(ent: str, com: str) -> list[str]:
    out: list[str] = []
    raw = Path(f"data/{ent}/interim/{com}_raw.parquet")
    report = Path(f"data/{ent}/output/{com}_kpi_report.parquet")

    if not raw.exists():
        return [f"\n[{ent}] no raw parquet — kedro step0_5 not done"]
    if not report.exists():
        return [f"\n[{ent}] no kpi_report.parquet — kedro step5 not done"]

    cfg = yaml.safe_load(Path(f"conf/{com}/parameters.yml").read_text(encoding="utf-8"))
    cols_cfg = cfg.get("columns", {})
    amt_col = cols_cfg.get("amount", "")
    ac_col = cols_cfg.get("account_code", "")
    ad_col = cols_cfg.get("account_desc", "")
    dn_col = cols_cfg.get("description", "")

    mappings = MAPPINGS.get(ent, {})
    if not mappings:
        return [f"\n[{ent}] no MAPPINGS — no comparison possible"]

    print(f"  reading raw...", flush=True)
    raw_df = pd.read_parquet(raw)
    print(f"  reading kpi_report...", flush=True)
    rep_df = pd.read_parquet(report)

    # Verify the kpi_report parquet has expected row-level columns
    needed = {"horizontal_id"}
    if not needed.issubset(rep_df.columns):
        return [f"\n[{ent}] kpi_report parquet has wrong schema (likely pivoted) — "
                f"re-run step5. Cols seen: {list(rep_df.columns)[:8]}..."]

    # Build sig per row in BOTH dataframes (assuming step5 preserves order/sig content)
    print(f"  building signatures...", flush=True)
    raw_df["_sig"] = _build_sig(raw_df, cfg)
    rep_df["_sig"] = _build_sig(rep_df, cfg)

    # Map raw_df with expected_h from manual label cols
    expected_h_series = pd.Series([None] * len(raw_df), index=raw_df.index, dtype="object")
    label_src_series = pd.Series([None] * len(raw_df), index=raw_df.index, dtype="object")
    for col, value_map in mappings.items():
        if col not in raw_df.columns:
            continue
        vals = raw_df[col].astype(str).fillna("")
        for v, h in value_map.items():
            mask = (expected_h_series.isna()) & (vals == v)
            expected_h_series.loc[mask] = h
            label_src_series.loc[mask] = f"{col}={v}"

    raw_df["_expected_h"] = expected_h_series
    raw_df["_label_src"] = label_src_series
    raw_df["_amt"] = pd.to_numeric(raw_df[amt_col], errors="coerce").fillna(0) if amt_col in raw_df.columns else 0

    # Build sig → our_h map from kpi_report
    sig_to_our_h = (rep_df.groupby("_sig")["horizontal_id"]
                          .agg(lambda s: s.value_counts().idxmax()))
    raw_df["_our_h"] = raw_df["_sig"].map(sig_to_our_h).fillna("(missing)")

    # Filter to rows with manual label
    labeled = raw_df.dropna(subset=["_expected_h"]).copy()
    n_label_rows = len(labeled)
    if n_label_rows == 0:
        return [f"\n[{ent}] no rows match MAPPINGS — check column names + values"]

    out.append("\n" + "=" * 100)
    out.append(f"[{ent}]  manual-labeled rows: {n_label_rows:,}/{len(raw_df):,} "
               f"({100*n_label_rows/len(raw_df):.1f}%)")
    out.append(f"  manual-labeled amount: {labeled['_amt'].sum():,.0f}")
    out.append("=" * 100)

    # ── SECTION A: overall agreement ──
    labeled["_match"] = labeled["_expected_h"] == labeled["_our_h"]
    n_match = int(labeled["_match"].sum())
    amt_match = float(labeled.loc[labeled["_match"], "_amt"].sum())
    amt_total_labeled = float(labeled["_amt"].sum())
    out.append(f"\n## A. AGREEMENT RATE ##")
    out.append(f"  rows match:   {n_match:,}/{n_label_rows:,} ({100*n_match/n_label_rows:.1f}%)")
    out.append(f"  amount match: {amt_match:,.0f}/{amt_total_labeled:,.0f} "
               f"({100*amt_match/amt_total_labeled if amt_total_labeled else 0:.1f}%)")

    # ── SECTION B: top disagreement buckets ──
    disagree = labeled[~labeled["_match"]].copy()
    out.append(f"\n## B. TOP DISAGREEMENT BUCKETS (expected_H → our_H) ##")
    if len(disagree) == 0:
        out.append("  ✅ no disagreements")
    else:
        bucket = (disagree.groupby(["_expected_h", "_our_h"])
                          .agg(_rows=("_amt", "count"), _amt=("_amt", "sum"))
                          .reset_index())
        bucket["_abs"] = bucket["_amt"].abs()
        bucket = bucket.sort_values("_abs", ascending=False).head(15)
        out.append(f"  {'expected':<18} {'our':<18} {'rows':>9} {'amount':>17}")
        for _, r in bucket.iterrows():
            out.append(f"  {str(r['_expected_h']):<18} {str(r['_our_h']):<18} "
                       f"{int(r['_rows']):>9,} {r['_amt']:>17,.0f}")

        # Detail: top sigs per bucket (show acct/desc context)
        out.append(f"\n  Top sigs per bucket (acct/desc context):")
        for _, r in bucket.head(8).iterrows():
            exp_h, our_h = r["_expected_h"], r["_our_h"]
            sub = disagree[(disagree["_expected_h"] == exp_h) & (disagree["_our_h"] == our_h)]
            sub_agg = (sub.groupby([ac_col, ad_col, dn_col], dropna=False)
                          .agg(_rows=("_amt", "count"), _amt=("_amt", "sum"))
                          .reset_index()
                          .sort_values("_amt", key=lambda s: s.abs(), ascending=False)
                          .head(5))
            out.append(f"\n    [{exp_h} → tagged as {our_h}]  (top 5)")
            for _, sig_r in sub_agg.iterrows():
                ad_v = str(sig_r.get(ad_col, ""))[:40]
                dn_v = str(sig_r.get(dn_col, ""))[:50]
                out.append(f"      acct='{ad_v}'  desc='{dn_v}'  rows={int(sig_r['_rows'])}  amt={sig_r['_amt']:,.0f}")

    # ── SECTION C: manual-labeled but tagged H_OTHER (instant-fix candidates) ──
    other_tagged = labeled[labeled["_our_h"] == "H_OTHER"]
    if len(other_tagged):
        out.append(f"\n## C. HIGH-VALUE FIXES: manual-labeled rows we tagged H_OTHER ##")
        out.append(f"  count: {len(other_tagged):,}  amount: {other_tagged['_amt'].sum():,.0f}")
        agg = (other_tagged.groupby([ad_col, dn_col, "_expected_h"], dropna=False)
                           .agg(_rows=("_amt", "count"), _amt=("_amt", "sum"))
                           .reset_index()
                           .sort_values("_amt", key=lambda s: s.abs(), ascending=False)
                           .head(15))
        for _, r in agg.iterrows():
            ad_v = str(r.get(ad_col, ""))[:35]
            dn_v = str(r.get(dn_col, ""))[:50]
            out.append(f"  → {r['_expected_h']:<15} acct='{ad_v}' desc='{dn_v}'  "
                       f"rows={int(r['_rows']):,} amt={r['_amt']:,.0f}")

    # ── SECTION D: LLM-teaching patterns ──
    # For each expected_H, find acct keywords that appear ONLY in rows manually labeled
    # to that H — these are strong signal patterns to add to LLM system prompt.
    out.append(f"\n## D. LLM-TEACHING PATTERNS (per expected_H) ##")
    out.append(f"  acct_desc keywords with >70% manual-label purity → add to LLM prompt")
    if ad_col in labeled.columns:
        for exp_h in sorted(labeled["_expected_h"].dropna().unique()):
            sub = labeled[labeled["_expected_h"] == exp_h]
            if len(sub) < 5:
                continue
            # Top acct_desc values for this expected_h
            ad_vc = sub[ad_col].astype(str).value_counts().head(10)
            # Purity check: for each acct_desc, how many rows total have it (across all expected_h)
            patterns = []
            for ad_val in ad_vc.index:
                if not ad_val or ad_val.lower() in EMPTY_VALUES:
                    continue
                total_rows_this_ad = labeled[labeled[ad_col].astype(str) == ad_val]
                if len(total_rows_this_ad) < 3:
                    continue
                purity = (total_rows_this_ad["_expected_h"] == exp_h).mean()
                if purity >= 0.70:
                    patterns.append((ad_val, int(ad_vc[ad_val]), purity))
            if patterns:
                out.append(f"\n  → {exp_h}:")
                for ad_val, n, pur in patterns[:8]:
                    out.append(f"      [{str(ad_val)[:60]:<60}]  {n:>5} rows  purity={pur*100:.0f}%")

    return out


def main():
    parser = argparse.ArgumentParser()
    g = parser.add_mutually_exclusive_group(required=True)
    g.add_argument("--entity", choices=list(ENTITIES))
    g.add_argument("--all", action="store_true")
    args = parser.parse_args()

    targets = list(ENTITIES.items()) if args.all else [(args.entity, ENTITIES[args.entity])]

    for ent, com in targets:
        print(f"\n>>> comparing {ent}...", flush=True)
        lines = compare_entity(ent, com)
        for line in lines:
            print(line, flush=True)
        out_path = Path(f"label_disagreement_{ent}.txt")
        out_path.write_text("\n".join(lines), encoding="utf-8")
        print(f"\n✓ wrote {out_path}")


if __name__ == "__main__":
    main()
