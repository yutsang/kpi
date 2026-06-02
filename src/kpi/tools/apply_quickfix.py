"""Apply manual corrections from the quickfix xlsx back to unique_projects/signatures.xlsx,
then re-run step4 + step5 automatically.

Usage:
  python -m kpi.tools.apply_quickfix --entity galaxy
  python -m kpi.tools.apply_quickfix --entity sjm

What it does:
  1. Reads <company>_quickfix.xlsx  (sheet 1_projects + sheet 2_signatures)
  2. For every row where `fix_vertical` is filled  → sets manual_vertical in unique_projects.xlsx
  3. For every row where `fix_horizontal` is filled → sets manual_horizontal in unique_signatures.xlsx
  4. Re-runs step4 (apply_tags) and step5 (build_report)
  5. Prints a summary of what was changed

Notes:
  - Existing manual tags that are NOT in the quickfix file are preserved.
  - `fix_vertical` must be a valid vertical_id (e.g. V_CONCERT). Invalid IDs are reported
    and skipped.
  - Run quickfix.py first to generate the xlsx, then fill it and run this script.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

import pandas as pd

from kpi.lib.conf import load_config, load_categories, path
from kpi.lib.io_setup import force_unbuffered_io

force_unbuffered_io()


def main():
    cfg = load_config()
    cats = load_categories()
    company = cfg["company"]["code"]
    interim = path(cfg, "interim_dir")
    out_dir = path(cfg, "output_dir")
    cols = cfg["columns"]
    proj_col = cols["project"]

    qf_path = out_dir / f"{company}_quickfix.xlsx"
    if not qf_path.exists():
        raise FileNotFoundError(
            f"Quickfix file not found: {qf_path}\n"
            f"Run first:  python -m kpi.tools.quickfix --entity "
            + (cfg["company"].get("alias") or company)
        )

    proj_xlsx = interim / f"{company}_unique_projects.xlsx"
    sig_xlsx = interim / f"{company}_unique_signatures.xlsx"
    for p in (proj_xlsx, sig_xlsx):
        if not p.exists():
            raise FileNotFoundError(f"Missing prerequisite: {p}")

    valid_v_ids = {v["id"] for v in cats["verticals"]}
    valid_h_ids = {h["id"] for h in cats["horizontals"] if h["id"] != "H_COUNT"}

    # ── Load quickfix sheets ──────────────────────────────────────────────────
    qf_proj = pd.read_excel(qf_path, sheet_name="1_projects")
    qf_sig = pd.read_excel(qf_path, sheet_name="2_signatures")

    # ── Apply vertical fixes to unique_projects.xlsx ──────────────────────────
    proj_df = pd.read_excel(proj_xlsx)
    if "manual_vertical" not in proj_df.columns:
        proj_df["manual_vertical"] = ""

    if proj_col not in qf_proj.columns or "fix_vertical" not in qf_proj.columns:
        print(f"  [skip] sheet 1_projects missing required columns ({proj_col}, fix_vertical)")
        v_applied = v_skipped = v_invalid = 0
    else:
        v_applied = v_skipped = v_invalid = 0
        for _, row in qf_proj.iterrows():
            fix = str(row.get("fix_vertical", "")).strip()
            if not fix or fix.lower() in ("nan", ""):
                continue
            proj_name = str(row.get(proj_col, "")).strip()
            if fix not in valid_v_ids:
                print(f"  [invalid] fix_vertical={fix!r} for project {proj_name!r} — not a valid vertical_id, skipped")
                v_invalid += 1
                continue
            mask = proj_df[proj_col].astype("string").fillna("").str.strip() == proj_name
            if mask.sum() == 0:
                print(f"  [not found] project {proj_name!r} not in unique_projects.xlsx — skipped")
                v_skipped += 1
                continue
            proj_df.loc[mask, "manual_vertical"] = fix
            v_applied += 1

    proj_df.to_excel(proj_xlsx, index=False)
    print(f"  vertical fixes: applied={v_applied}  skipped={v_skipped}  invalid={v_invalid}", flush=True)

    # ── Apply horizontal fixes to unique_signatures.xlsx ─────────────────────
    sig_df = pd.read_excel(sig_xlsx)
    if "manual_horizontal" not in sig_df.columns:
        sig_df["manual_horizontal"] = ""

    # Build a signature key for matching: account_code | account_desc | desc_norm
    def _make_sig(r) -> str:
        return (
            str(r.get("account_code", "")).strip() + "|"
            + str(r.get("account_desc", "")).strip() + "|"
            + str(r.get("desc_norm", "")).strip()
        )

    sig_df["_sig_key"] = sig_df.apply(_make_sig, axis=1)
    sig_key_index = {k: i for i, k in enumerate(sig_df["_sig_key"])}

    h_applied = h_skipped = h_invalid = 0
    if "fix_horizontal" not in qf_sig.columns:
        print(f"  [skip] sheet 2_signatures missing fix_horizontal column")
    else:
        for _, row in qf_sig.iterrows():
            fix = str(row.get("fix_horizontal", "")).strip()
            if not fix or fix.lower() in ("nan", ""):
                continue
            if fix not in valid_h_ids:
                print(f"  [invalid] fix_horizontal={fix!r} — not a valid horizontal_id, skipped")
                h_invalid += 1
                continue
            # Try to match by signature key first
            row_sig_key = _make_sig(row)
            if row_sig_key in sig_key_index:
                idx = sig_key_index[row_sig_key]
                sig_df.at[idx, "manual_horizontal"] = fix
                h_applied += 1
            else:
                # Fallback: match by account_code + account_desc only
                acct = str(row.get("account_code", "")).strip()
                adesc = str(row.get("account_desc", "")).strip()
                mask = (
                    sig_df["account_code"].astype("string").fillna("").str.strip() == acct
                ) & (
                    sig_df["account_desc"].astype("string").fillna("").str.strip() == adesc
                )
                if mask.sum() > 0:
                    sig_df.loc[mask, "manual_horizontal"] = fix
                    h_applied += mask.sum()
                else:
                    h_skipped += 1

    sig_df = sig_df.drop(columns=["_sig_key"], errors="ignore")
    sig_df.to_excel(sig_xlsx, index=False)
    print(f"  horizontal fixes: applied={h_applied}  skipped={h_skipped}  invalid={h_invalid}", flush=True)

    total_changes = v_applied + h_applied
    if total_changes == 0:
        print("\n  No changes applied — nothing to re-run.", flush=True)
        print("  Tip: fill `fix_vertical` / `fix_horizontal` in the quickfix xlsx first.")
        return

    # ── Re-run step4 + step5 ─────────────────────────────────────────────────
    print(f"\n  {total_changes} corrections applied. Re-running step4 + step5 …", flush=True)
    os.environ["KPI_ENTITY"] = company

    from kpi.pipelines.step4_apply_tags import _logic as step4
    from kpi.pipelines.step5_build_report import _logic as step5

    print("\n── step4: apply tags ──", flush=True)
    step4.main()
    print("\n── step5: build report ──", flush=True)
    step5.main()
    print(f"\nDone. Check data/{cfg['company'].get('folder', company)}/output/ for updated report.", flush=True)


if __name__ == "__main__":
    main()
