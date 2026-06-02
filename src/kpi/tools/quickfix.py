"""Generate a focused manual-correction workbook for one entity.

Usage:
  python -m kpi.tools.quickfix --entity galaxy
  python -m kpi.tools.quickfix --entity sjm

Outputs:
  data/<alias>/output/<company>_quickfix.xlsx

Sheets
------
  0_zero_analysis   — every vertical's current amount + row count; 0s shown clearly
  1_projects        — all unique projects sorted by |amount|; fill `fix_vertical` column
  2_signatures      — all unique signatures sorted by |amount|; fill `fix_horizontal` column
  3_untagged_rows   — rows still sitting as (未分類) after step4, sorted by |amount|

Workflow
--------
  1. Run this script  →  open <company>_quickfix.xlsx
  2. In sheet 1_projects:  fill `fix_vertical` for any wrong/missing vertical tag
     (use a vertical_id from categories.yml, e.g. V_CONCERT, V_RESTAURANT …)
  3. In sheet 2_signatures: fill `fix_horizontal` for any wrong/missing horizontal tag
     (use a horizontal_id, e.g. H_CONSTRUCTION, H_FNB, H_LABOR …)
  4. Save the file, then run:
       python -m kpi.tools.apply_quickfix --entity <entity>
  Corrections are written to unique_projects.xlsx and unique_signatures.xlsx,
  then step4 + step5 re-run automatically.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

from kpi.lib.conf import load_config, load_categories, path
from kpi.lib.io_setup import force_unbuffered_io

force_unbuffered_io()


def main(entity_or_alias: str | None = None):
    cfg = load_config(entity_or_alias)
    cats = load_categories()
    company = cfg["company"]["code"]
    interim = path(cfg, "interim_dir")
    out_dir = path(cfg, "output_dir")
    cols = cfg["columns"]

    proj_xlsx = interim / f"{company}_unique_projects.xlsx"
    sig_xlsx = interim / f"{company}_unique_signatures.xlsx"
    tagged_path = interim / f"{company}_tagged_rows.parquet"

    missing = [p for p in (proj_xlsx, sig_xlsx) if not p.exists()]
    if missing:
        raise FileNotFoundError(f"Run step1–step3 first. Missing: {[p.name for p in missing]}")

    proj_df = pd.read_excel(proj_xlsx)
    sig_df = pd.read_excel(sig_xlsx)
    df_tagged = pd.read_parquet(tagged_path) if tagged_path.exists() else None

    # --- reference lists ---
    v_list = [(v["id"], v["label"]) for v in cats["verticals"]]
    h_list = [(h["id"], h["label"]) for h in cats["horizontals"] if h["id"] != "H_COUNT"]

    out_path = out_dir / f"{company}_quickfix.xlsx"
    print(f"Building {out_path.name} …", flush=True)

    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        wb = writer.book

        # ── formats ──────────────────────────────────────────────────────────
        fmt_header = wb.add_format({"bold": True, "bg_color": "#2F5496", "font_color": "white",
                                    "border": 1, "text_wrap": True})
        fmt_zero = wb.add_format({"bg_color": "#FFD7D7", "bold": True})
        fmt_low  = wb.add_format({"bg_color": "#FFF2CC"})
        fmt_fix  = wb.add_format({"bg_color": "#E2EFDA", "bold": True,
                                   "border": 2, "border_color": "#375623"})
        fmt_note = wb.add_format({"italic": True, "font_color": "#7F7F7F", "text_wrap": True})

        def _write_df(df: pd.DataFrame, sheet: str, fix_cols: list[str] | None = None):
            # Deduplicate columns before writing (duplicate names → pandas returns DataFrame
            # instead of Series on column access, breaking .str operations downstream)
            if df.columns.duplicated().any():
                df = df.loc[:, ~df.columns.duplicated()]
            df.to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.sheets[sheet]
            for col_idx, col_name in enumerate(df.columns):
                width = max(12, min(45, len(str(col_name)) + 2))
                try:
                    col_data = df[col_name]
                    if isinstance(col_data, pd.DataFrame):
                        col_data = col_data.iloc[:, 0]
                    sample = col_data.astype(str).str.len().quantile(0.9) if len(df) > 0 else 10
                except Exception:
                    sample = 15
                width = max(width, min(50, int(sample) + 2))
                if fix_cols and col_name in fix_cols:
                    ws.set_column(col_idx, col_idx, width, fmt_fix)
                else:
                    ws.set_column(col_idx, col_idx, width)
            ws.freeze_panes(1, 0)
            return ws

        # ── Sheet 0: Zero analysis ────────────────────────────────────────────
        v_amounts: dict[str, float] = {}
        v_counts: dict[str, int] = {}
        if df_tagged is not None:
            df_tagged["_amt"] = pd.to_numeric(df_tagged[cols["amount"]], errors="coerce").fillna(0)
            grp = df_tagged.groupby("vertical_label", observed=False)["_amt"]
            v_amounts = grp.sum().to_dict()
            v_counts = df_tagged.groupby("vertical_label", observed=False).size().to_dict()

        zero_rows = []
        for vid, vlabel in v_list:
            amt = v_amounts.get(vlabel, 0)
            cnt = v_counts.get(vlabel, 0)
            zero_rows.append({
                "vertical_id": vid,
                "vertical_label": vlabel,
                "total_amount": amt,
                "row_count": cnt,
                "flag": "⚠ ZERO" if amt == 0 else ("low" if cnt < 5 else ""),
            })
        zero_df = pd.DataFrame(zero_rows).sort_values("total_amount")

        # Also show horizontal totals
        h_amounts: dict[str, float] = {}
        h_counts: dict[str, int] = {}
        if df_tagged is not None:
            hgrp = df_tagged.groupby("horizontal_label", observed=False)["_amt"]
            h_amounts = hgrp.sum().to_dict()
            h_counts = df_tagged.groupby("horizontal_label", observed=False).size().to_dict()

        h_rows = []
        for hid, hlabel in h_list:
            amt = h_amounts.get(hlabel, 0)
            cnt = h_counts.get(hlabel, 0)
            h_rows.append({
                "horizontal_id": hid,
                "horizontal_label": hlabel,
                "total_amount": amt,
                "row_count": cnt,
                "flag": "⚠ ZERO" if amt == 0 else "",
            })
        h_df = pd.DataFrame(h_rows).sort_values("total_amount")

        # Write combined zero analysis
        zero_df.to_excel(writer, sheet_name="0_zero_analysis", index=False, startrow=0)
        h_df.to_excel(writer, sheet_name="0_zero_analysis", index=False,
                      startrow=len(zero_df) + 3)
        ws0 = writer.sheets["0_zero_analysis"]
        ws0.set_column(0, 5, 22)
        ws0.freeze_panes(1, 0)
        # Highlight zero rows
        for row_idx, row in enumerate(zero_rows, start=1):
            if row["flag"] == "⚠ ZERO":
                ws0.set_row(row_idx, None, fmt_zero)
        print(f"  Zero verticals: {sum(1 for r in zero_rows if r['flag']=='⚠ ZERO')}"
              f"  Zero horizontals: {sum(1 for r in h_rows if r['flag']=='⚠ ZERO')}",
              flush=True)

        # ── Sheet 1: Projects ─────────────────────────────────────────────────
        proj_col = cols["project"]
        proj_show_cols = [
            proj_col, "total_amount", "row_count",
            "primary_ng",
            "llm_vertical", "llm_vertical_label", "llm_confidence", "llm_reasoning",
            "llm_alt_candidates",
        ]
        proj_show_cols = [c for c in proj_show_cols if c in proj_df.columns]
        proj_out = proj_df[proj_show_cols].copy()
        proj_out["fix_vertical"] = proj_df.get("manual_vertical", "")  # carry forward existing manual
        proj_out["_abs"] = pd.to_numeric(proj_out.get("total_amount", 0), errors="coerce").abs().fillna(0)
        proj_out = proj_out.sort_values("_abs", ascending=False).drop(columns="_abs")

        _write_df(proj_out, "1_projects", fix_cols=["fix_vertical"])
        print(f"  Projects: {len(proj_out)}", flush=True)

        # ── Sheet 2: Signatures ───────────────────────────────────────────────
        sig_show_cols = [
            "account_code", "account_desc", "desc_norm", "desc_samples", "project_samples",
            "total_amount", "row_count",
            "capex_opex_dominant", "capex_opex_dominant_pct",
            "llm_horizontal", "llm_horizontal_label", "llm_confidence", "llm_reasoning",
        ]
        sig_show_cols = [c for c in sig_show_cols if c in sig_df.columns]
        sig_out = sig_df[sig_show_cols].copy()
        sig_out["fix_horizontal"] = sig_df.get("manual_horizontal", "")
        sig_out["_abs"] = pd.to_numeric(sig_out.get("total_amount", 0), errors="coerce").abs().fillna(0)
        sig_out = sig_out.sort_values("_abs", ascending=False).drop(columns="_abs")

        _write_df(sig_out, "2_signatures", fix_cols=["fix_horizontal"])
        print(f"  Signatures: {len(sig_out)}", flush=True)

        # ── Sheet 3: Untagged rows ────────────────────────────────────────────
        if df_tagged is not None:
            untagged_mask = (
                df_tagged["vertical_label"].fillna("").isin(["", "(未分類)"])
                | df_tagged["horizontal_label"].fillna("").isin(["", "(未分類)"])
            )
            untagged = df_tagged[untagged_mask].copy()
            row_cols = [
                cols.get("period"), cols.get("posting_year"),
                cols.get("project"), cols.get("ng11_category"), cols.get("capex_opex"),
                cols.get("amount"), cols.get("account_code"), cols.get("account_desc"),
                cols.get("description"), cols.get("vendor"),
                "vertical_label", "horizontal_label", "final_capex_opex", "row_type",
            ]
            row_cols = [c for c in row_cols if c and c in untagged.columns]
            untagged = untagged.sort_values("_amt", ascending=False).head(1000)[row_cols]
            _write_df(untagged, "3_untagged_rows")
            print(f"  Untagged rows (top 1000): {len(untagged)}", flush=True)
        else:
            pd.DataFrame([{"note": "Run step4 first to see untagged rows"}]).to_excel(
                writer, sheet_name="3_untagged_rows", index=False
            )

        # ── Sheet 4: Reference — valid IDs ───────────────────────────────────
        ref = pd.DataFrame(
            [{"type": "vertical", "id": vid, "label": vlabel} for vid, vlabel in v_list]
            + [{"type": "horizontal", "id": hid, "label": hlabel} for hid, hlabel in h_list]
        )
        _write_df(ref, "4_valid_ids")

    print(f"\nWrote: {out_path}", flush=True)
    print("\nNext steps:")
    print("  1. Open the xlsx and check sheet 0_zero_analysis — which rows are ⚠ ZERO?")
    print("  2. Sheet 1_projects — fill `fix_vertical` (use IDs from sheet 4_valid_ids)")
    print("  3. Sheet 2_signatures — fill `fix_horizontal`")
    print("  4. Save, then run:  python -m kpi.tools.apply_quickfix --entity "
          + (cfg["company"].get("alias") or company))


def run_all():
    """Run quickfix for every configured entity that has data (unique_projects.xlsx exists)."""
    from kpi.lib.conf import load_master

    master = load_master()
    companies = master.get("companies") or {}
    results = []

    for key, ent in companies.items():
        if not ent or not ent.get("name"):
            continue
        alias = (ent.get("alias") or key).strip()
        folder = alias
        interim_dir = (
            Path(__file__).resolve().parents[3] / "data" / folder / "interim"
        )
        proj_check = interim_dir / f"{key}_unique_projects.xlsx"
        if not proj_check.exists():
            print(f"  [skip] {key} ({alias}) — no unique_projects.xlsx yet", flush=True)
            results.append((key, alias, "skipped"))
            continue

        print(f"\n{'='*60}\n  {key}  ({alias})\n{'='*60}", flush=True)
        try:
            main(entity_or_alias=key)   # pass directly, no env var needed
            results.append((key, alias, "ok"))
        except Exception as e:
            print(f"  ERROR for {key}: {e}", flush=True)
            results.append((key, alias, f"error: {e}"))

    print("\n── Summary ─────────────────────────────────────────────────")
    for key, alias, status in results:
        print(f"  {key:14s} ({alias:8s})  {status}")


if __name__ == "__main__":
    import argparse as _ap
    _parser = _ap.ArgumentParser(add_help=False)
    _parser.add_argument("--all", action="store_true")
    _parser.add_argument("--entity", default=None)
    _args, _ = _parser.parse_known_args()

    if _args.all:
        run_all()
    else:
        main(entity_or_alias=_args.entity)
