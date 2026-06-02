"""Offsets pipeline — detect offset / reversal entries to support audit sampling.

Reads <entity>_tagged_rows.parquet (step 4 output) and outputs
<entity>_offsets_review.xlsx where each row carries an `offset_tag` indicating which
detection layer caught it. Auditors can filter/sort by `offset_tag` + `pair_id` to
identify paired entries that should be sampled together (or skipped because they
net to zero).

Detection layers (high → low confidence):
  1. strict_pair        — same project + account + vendor + |amount|, opposite
                          sign, posting_date within `pair_window_days`
  2. reversal_keyword   — description contains a reversal keyword
                          (沖回 / 沖銷 / 取消 / reversal / cancel / refund / void / write-off ...)
  3. aggregate_reversal — single negative whose |amount| ≈ sum of multiple positives
                          in the same project + account, within `aggregate_window_days`
  4. net_zero_signature — signature whose gross > 0 but net rolls to ≈0 (after
                          all positives + negatives) — economically nothing happened

Rows untagged after layers 1-4 are NOT included in the output xlsx (they have no
offset/reversal pattern detected — auditor should treat them as standalone entries).

Generic across all 6 entities — uses the entity's column mapping from cfg["columns"].
"""
from __future__ import annotations

import pandas as pd

from kpi.lib.conf import load_config, path
from kpi.lib.io_setup import force_unbuffered_io

force_unbuffered_io()


DEFAULT_OPTS = {
    "pair_window_days":         90,    # max days between a +/- pair to consider them paired
    "amount_tolerance":         0.01,  # absolute MOP tolerance for amount equality
    "aggregate_window_days":    365,   # max days for many-positives -> 1-negative cluster
    "aggregate_tolerance_pct":  0.01,  # 1% relative tolerance for aggregate sum match
    "min_aggregate_positives":  2,     # need at least this many positives in a cluster
    "reversal_keywords": [
        # Chinese
        "沖回", "反沖", "沖銷", "沖抵", "取消", "撤銷", "退款", "退回",
        # English (case-insensitive matched)
        "reversal", "reverse", "cancel", "cancellation", "cancelled", "canceled",
        "refund", "refunded", "void", "voided",
        "write-off", "writeoff", "write off", "waiver", "waived",
        "accrual reversal", "reverse accrual",
    ],
}


def _resolve_opts(cfg: dict) -> dict:
    """Merge user `offsets:` block from parameters.yml (if any) over defaults."""
    user = cfg.get("offsets") or (cfg.get("_master") or {}).get("offsets") or {}
    return {**DEFAULT_OPTS, **(user or {})}


def main():
    cfg = load_config()
    company = cfg["company"]["code"]
    interim = path(cfg, "interim_dir")
    output_dir = path(cfg, "output_dir")
    cols = cfg["columns"]
    opts = _resolve_opts(cfg)

    parquet = interim / f"{company}_tagged_rows.parquet"
    if not parquet.exists():
        raise FileNotFoundError(f"Run step 4 first. Missing: {parquet}")

    df = pd.read_parquet(parquet)
    print(f"Loaded {len(df):,} rows from {parquet.name}", flush=True)

    amount_col = cols["amount"]
    project_col = cols["project"]
    account_col = cols["account_code"]
    vendor_col = cols.get("vendor", "")
    desc_col = cols.get("description", "")
    date_col = cols.get("posting_date", "")

    # Normalize working columns (don't mutate originals; prefix with _)
    df["_amt"] = pd.to_numeric(df[amount_col], errors="coerce").fillna(0.0)
    df["_abs"] = df["_amt"].abs()
    if date_col and date_col in df.columns:
        df["_date"] = pd.to_datetime(df[date_col], errors="coerce")
    else:
        df["_date"] = pd.NaT
    df["_proj"] = df[project_col].astype("string").fillna("").str.strip() if project_col in df.columns else ""
    df["_acc"]  = df[account_col].astype("string").fillna("").str.strip() if account_col in df.columns else ""
    if vendor_col and vendor_col in df.columns:
        df["_vendor"] = df[vendor_col].astype("string").fillna("").str.strip()
        has_vendor = True
    else:
        df["_vendor"] = ""
        has_vendor = False
    if desc_col and desc_col in df.columns:
        df["_desc"] = df[desc_col].astype("string").fillna("").str.strip()
    else:
        df["_desc"] = ""

    df["offset_tag"] = ""
    df["pair_id"] = ""

    # ============ LAYER 1: strict pair ============
    print("\nLayer 1: strict pairs (same project + account"
          + (" + vendor" if has_vendor else "")
          + ", |pos|≈|neg|, within window) ...", flush=True)
    pair_counter = 0
    grp_keys = ["_proj", "_acc"] + (["_vendor"] if has_vendor else []) + ["_abs_r"]
    df["_abs_r"] = df["_abs"].round(2)

    pair_window = pd.Timedelta(days=opts["pair_window_days"])
    tol = float(opts["amount_tolerance"])

    for _, sub in df.groupby(grp_keys, sort=False):
        if len(sub) < 2:
            continue
        pos = sub[sub["_amt"] > tol]
        neg = sub[sub["_amt"] < -tol]
        if len(pos) == 0 or len(neg) == 0:
            continue
        # Greedy: for each pos, find nearest-date neg with matching |amount| (already
        # ensured by group key) within window. Each neg can be claimed at most once.
        used_neg: set = set()
        pos_sorted = pos.sort_values("_date")
        for p_idx in pos_sorted.index:
            p_date = pos_sorted.loc[p_idx, "_date"]
            best = None
            best_dist = None
            for n_idx in neg.index:
                if n_idx in used_neg:
                    continue
                n_date = neg.loc[n_idx, "_date"]
                if pd.notna(p_date) and pd.notna(n_date):
                    delta = abs(p_date - n_date)
                    if delta > pair_window:
                        continue
                    dist = delta.days
                else:
                    dist = 10**9  # missing date — accept but heavily de-prefer
                if best_dist is None or dist < best_dist:
                    best = n_idx
                    best_dist = dist
            if best is not None:
                pair_counter += 1
                pid = f"P{pair_counter:06d}"
                df.at[p_idx, "offset_tag"] = "strict_pair"
                df.at[p_idx, "pair_id"] = pid
                df.at[best, "offset_tag"] = "strict_pair"
                df.at[best, "pair_id"] = pid
                used_neg.add(best)

    n_strict = int((df["offset_tag"] == "strict_pair").sum())
    amt_strict = float(df.loc[df["offset_tag"] == "strict_pair", "_abs"].sum())
    print(f"  found {n_strict:,} rows in {pair_counter:,} pairs  |amount|={amt_strict:,.0f}", flush=True)

    # ============ LAYER 2: reversal keyword ============
    print("\nLayer 2: reversal keyword in description ...", flush=True)
    kws = [k for k in opts["reversal_keywords"] if k]
    if kws:
        # Escape regex special chars by using re.escape via pandas
        import re
        pattern = "|".join(re.escape(k) for k in kws)
        has_kw = df["_desc"].str.contains(pattern, case=False, na=False, regex=True)
        new_kw = has_kw & (df["offset_tag"] == "")
        df.loc[new_kw, "offset_tag"] = "reversal_keyword"
        n_kw = int(new_kw.sum())
        amt_kw = float(df.loc[new_kw, "_abs"].sum())
    else:
        n_kw = 0
        amt_kw = 0.0
    print(f"  flagged {n_kw:,} rows  |amount|={amt_kw:,.0f}", flush=True)

    # ============ LAYER 3: aggregate reversal ============
    print("\nLayer 3: aggregate reversals (N positives offset by 1 negative) ...", flush=True)
    cluster_counter = 0
    agg_window = pd.Timedelta(days=opts["aggregate_window_days"])
    agg_tol_pct = float(opts["aggregate_tolerance_pct"])
    min_pos = int(opts["min_aggregate_positives"])

    # Only consider rows still untagged
    untagged_mask = df["offset_tag"] == ""
    df_u = df[untagged_mask]

    for _, sub in df_u.groupby(["_proj", "_acc"], sort=False):
        if len(sub) < (min_pos + 1):
            continue
        negs = sub[sub["_amt"] < -tol]
        if len(negs) == 0:
            continue
        for n_idx in negs.index:
            if df.at[n_idx, "offset_tag"] != "":
                continue
            neg_amt = abs(float(sub.loc[n_idx, "_amt"]))
            n_date = sub.loc[n_idx, "_date"]
            cands = sub[(sub["_amt"] > tol) & (sub["offset_tag"] == "")]
            if pd.notna(n_date):
                cands = cands[
                    (cands["_date"].notna())
                    & (cands["_date"] <= n_date)
                    & ((n_date - cands["_date"]) <= agg_window)
                ]
            if len(cands) < min_pos:
                continue
            pos_sum = float(cands["_amt"].sum())
            denom = max(neg_amt, 1.0)
            if abs(pos_sum - neg_amt) / denom <= agg_tol_pct:
                cluster_counter += 1
                cid = f"A{cluster_counter:06d}"
                for c_idx in cands.index:
                    df.at[c_idx, "offset_tag"] = "aggregate_reversal"
                    df.at[c_idx, "pair_id"] = cid
                df.at[n_idx, "offset_tag"] = "aggregate_reversal"
                df.at[n_idx, "pair_id"] = cid

    n_agg = int((df["offset_tag"] == "aggregate_reversal").sum())
    amt_agg = float(df.loc[df["offset_tag"] == "aggregate_reversal", "_abs"].sum())
    print(f"  found {n_agg:,} rows in {cluster_counter:,} clusters  |amount|={amt_agg:,.0f}", flush=True)

    # ============ LAYER 4: net-zero signatures ============
    print("\nLayer 4: net-zero signatures (gross > 0 but net rolls to ≈0) ...", flush=True)
    if "signature" in df.columns:
        sig_agg = df.groupby("signature").agg(
            gross=("_abs", "sum"),
            net=("_amt", "sum"),
        )
        net_zero_sigs = sig_agg[
            (sig_agg["gross"] > tol) & (sig_agg["net"].abs() < tol)
        ].index
        nz_mask = df["signature"].isin(net_zero_sigs) & (df["offset_tag"] == "")
        df.loc[nz_mask, "offset_tag"] = "net_zero_signature"
        # Use signature as the pair_id for net_zero so user can filter
        df.loc[nz_mask, "pair_id"] = "S:" + df.loc[nz_mask, "signature"].astype(str)
        n_nz = int(nz_mask.sum())
        amt_nz = float(df.loc[nz_mask, "_abs"].sum())
        n_nz_groups = int(len(net_zero_sigs))
    else:
        n_nz = 0
        amt_nz = 0.0
        n_nz_groups = 0
    print(f"  flagged {n_nz:,} rows in {n_nz_groups:,} signatures  |amount|={amt_nz:,.0f}", flush=True)

    # ============ Output xlsx ============
    print("\nWriting xlsx ...", flush=True)
    out_path = output_dir / f"{company}_offsets_review.xlsx"

    keep_cols = []
    for c in [
        "offset_tag",
        "pair_id",
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
        "vertical_label",
        "horizontal_label",
    ]:
        if c and c in df.columns and c not in keep_cols:
            keep_cols.append(c)

    flagged = df[df["offset_tag"] != ""][keep_cols].copy()
    # Sort: tag first, then pair_id, then amount desc within pair
    if cols.get("amount", "") in flagged.columns:
        flagged = flagged.sort_values(
            ["offset_tag", "pair_id", cols["amount"]],
            ascending=[True, True, False],
        )
    else:
        flagged = flagged.sort_values(["offset_tag", "pair_id"])

    n_flag = len(flagged)
    n_unflag = len(df) - n_flag
    amt_flag = float(df.loc[df["offset_tag"] != "", "_abs"].sum())
    amt_unflag = float(df.loc[df["offset_tag"] == "", "_abs"].sum())
    total_abs = amt_flag + amt_unflag

    # Excel hard limit ~1,048,576 rows per sheet — Galaxy 624k flagged is well under,
    # but defensively cap.
    EXCEL_ROW_LIMIT = 1_000_000
    if n_flag > EXCEL_ROW_LIMIT:
        print(f"  ⚠ flagged rows {n_flag:,} exceeds Excel limit; truncating to {EXCEL_ROW_LIMIT:,}", flush=True)
        flagged = flagged.head(EXCEL_ROW_LIMIT)

    summary_rows = [
        {"layer": "1. strict_pair",
         "rows": n_strict, "groups": pair_counter, "amount_abs": int(amt_strict),
         "pct_of_total_amt": round(amt_strict / total_abs * 100, 2) if total_abs > 0 else 0,
         "description": "same project + account"
                        + (" + vendor" if has_vendor else "")
                        + ", |pos|=|neg| within window — pure cancellations / re-bookings"},
        {"layer": "2. reversal_keyword",
         "rows": n_kw, "groups": "", "amount_abs": int(amt_kw),
         "pct_of_total_amt": round(amt_kw / total_abs * 100, 2) if total_abs > 0 else 0,
         "description": "description contains 沖回 / 沖銷 / 取消 / reversal / cancel / refund / void / etc."},
        {"layer": "3. aggregate_reversal",
         "rows": n_agg, "groups": cluster_counter, "amount_abs": int(amt_agg),
         "pct_of_total_amt": round(amt_agg / total_abs * 100, 2) if total_abs > 0 else 0,
         "description": "N positives in same project+account offset by 1 net-equal negative — accrual + reversal pattern"},
        {"layer": "4. net_zero_signature",
         "rows": n_nz, "groups": n_nz_groups, "amount_abs": int(amt_nz),
         "pct_of_total_amt": round(amt_nz / total_abs * 100, 2) if total_abs > 0 else 0,
         "description": "signature whose gross > 0 but net rolls to ~0 — economically nothing happened"},
        {"layer": "TOTAL FLAGGED",
         "rows": n_flag, "groups": "", "amount_abs": int(amt_flag),
         "pct_of_total_amt": round(amt_flag / total_abs * 100, 2) if total_abs > 0 else 0,
         "description": ""},
        {"layer": "(unflagged — no offset detected)",
         "rows": n_unflag, "groups": "", "amount_abs": int(amt_unflag),
         "pct_of_total_amt": round(amt_unflag / total_abs * 100, 2) if total_abs > 0 else 0,
         "description": "treat as standalone entries during sampling"},
    ]
    summary_df = pd.DataFrame(summary_rows)

    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        summary_df.to_excel(writer, sheet_name="0_summary", index=False)
        flagged.to_excel(writer, sheet_name="1_flagged_rows", index=False)

    size_mb = out_path.stat().st_size / 1e6
    print(f"\nWrote {out_path}  ({size_mb:.1f} MB, {n_flag:,} flagged rows)", flush=True)
    print(f"\nSampling workflow:", flush=True)
    print(f"  1. Pick a sample row from main tagged_rows xlsx (or PowerBI fact)", flush=True)
    print(f"  2. Match by (posting_date, amount, project, account, description) in this file's 1_flagged_rows", flush=True)
    print(f"  3. If found: filter by `pair_id` to see the offsetting mate(s) — sample them together", flush=True)
    print(f"  4. If absent: row is NOT part of any offset/reversal — treat as standalone", flush=True)
    print(f"\n  Tag legend:", flush=True)
    print(f"    strict_pair         → 1 pos + 1 neg (matching |amount|, ≤{opts['pair_window_days']}d apart)", flush=True)
    print(f"    reversal_keyword    → description has 沖回/reversal/cancel/etc.", flush=True)
    print(f"    aggregate_reversal  → N positives netted by 1 negative", flush=True)
    print(f"    net_zero_signature  → entire signature nets to zero (gross != 0)", flush=True)
