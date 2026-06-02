"""Sample signatures across all companies for AI review.

Usage:
  python -m kpi.tools.sample_sigs          # all companies with data
  python -m kpi.tools.sample_sigs --entity galaxy  # one company only

Output:
  data/audit/sig_sample_<date>.xlsx

The file has one sheet per company, each with ~60 rows:
  - 20 highest-amount signatures (covering most of the money)
  - 20 H_OTHER signatures by amount  (LLM didn't know what to pick)
  - 20 low-confidence (<0.7) by amount (LLM was unsure)

Paste the sheet (or describe rows) to the AI for prompt diagnosis.
"""
from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[3]
AUDIT_DIR = ROOT / "data" / "audit"


def _sample_one(sig_df: pd.DataFrame, company: str, alias: str, n: int = 20) -> pd.DataFrame:
    df = sig_df.copy()
    df["_abs"] = pd.to_numeric(df.get("total_amount", 0), errors="coerce").abs().fillna(0)

    keep_cols = [
        "account_code", "account_desc", "desc_norm", "desc_samples",
        "total_amount", "row_count",
        "capex_opex_dominant", "capex_opex_dominant_pct",
        "llm_horizontal", "llm_horizontal_label",
        "llm_confidence", "llm_reasoning",
        "llm_status",
    ]
    keep_cols = [c for c in keep_cols if c in df.columns]

    parts = []

    # Top N by amount
    top = df.sort_values("_abs", ascending=False).head(n).copy()
    top["_section"] = "top_amount"
    parts.append(top)

    # H_OTHER
    if "llm_horizontal" in df.columns:
        h_other = df[df["llm_horizontal"].astype("string").fillna("") == "H_OTHER"]
        h_other = h_other.sort_values("_abs", ascending=False).head(n).copy()
        h_other["_section"] = "H_OTHER"
        parts.append(h_other)

    # Low confidence
    if "llm_confidence" in df.columns:
        low = df[pd.to_numeric(df["llm_confidence"], errors="coerce") < 0.7]
        low = low.sort_values("_abs", ascending=False).head(n).copy()
        low["_section"] = "low_conf"
        parts.append(low)

    out = (
        pd.concat(parts)
        .drop_duplicates(subset=["account_code", "account_desc", "desc_norm"]
                         if all(c in df.columns for c in ["account_code", "account_desc", "desc_norm"])
                         else None)
        .drop(columns=["_abs"], errors="ignore")
    )
    out.insert(0, "company", alias)
    out.insert(1, "entity_key", company)
    final_cols = ["company", "entity_key", "_section"] + [c for c in keep_cols if c in out.columns]
    return out[[c for c in final_cols if c in out.columns]]


def run(entity_or_alias: str | None = None):
    from kpi.lib.conf import load_master, load_config, path

    AUDIT_DIR.mkdir(parents=True, exist_ok=True)

    if entity_or_alias:
        # Single entity
        cfg = load_config(entity_or_alias)
        company = cfg["company"]["code"]
        alias = cfg["company"].get("alias") or company
        interim = path(cfg, "interim_dir")
        sig_xlsx = interim / f"{company}_unique_signatures.xlsx"
        if not sig_xlsx.exists():
            raise FileNotFoundError(f"Run step1+step3 first. Missing: {sig_xlsx}")
        sig_df = pd.read_excel(sig_xlsx)
        entities = [(company, alias, sig_df)]
    else:
        # All configured entities
        master = load_master()
        entities = []
        for key, ent in (master.get("companies") or {}).items():
            if not ent or not ent.get("name"):
                continue
            alias = (ent.get("alias") or key).strip()
            folder = alias
            interim_dir = ROOT / "data" / folder / "interim"
            sig_xlsx = interim_dir / f"{key}_unique_signatures.xlsx"
            if not sig_xlsx.exists():
                print(f"  [skip] {key} — no unique_signatures.xlsx yet")
                continue
            sig_df = pd.read_excel(sig_xlsx)
            entities.append((key, alias, sig_df))

    if not entities:
        print("No entities with signature data found.")
        return

    fname = f"sig_sample_{date.today().isoformat()}.xlsx"
    out_path = AUDIT_DIR / fname

    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        wb = writer.book
        fmt_h = wb.add_format({"bold": True, "bg_color": "#2F5496", "font_color": "white"})
        fmt_other = wb.add_format({"bg_color": "#FFF2CC"})
        fmt_low   = wb.add_format({"bg_color": "#FFD7D7"})

        all_parts = []
        for company, alias, sig_df in entities:
            sample = _sample_one(sig_df, company, alias)
            all_parts.append(sample)

            # Per-company sheet
            sheet_name = alias[:31]
            sample.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]
            ws.set_column(0, len(sample.columns) - 1, 20)
            ws.freeze_panes(1, 0)
            # Colour rows by section
            for row_idx, sec in enumerate(sample.get("_section", []), start=1):
                if str(sec) == "H_OTHER":
                    ws.set_row(row_idx, None, fmt_other)
                elif str(sec) == "low_conf":
                    ws.set_row(row_idx, None, fmt_low)

            print(f"  {alias:12s}: {len(sample)} rows sampled from {len(sig_df)} signatures")

        # Combined sheet (all companies together)
        if len(all_parts) > 1:
            combined = pd.concat(all_parts, ignore_index=True)
            combined.to_excel(writer, sheet_name="ALL_combined", index=False)
            ws_all = writer.sheets["ALL_combined"]
            ws_all.set_column(0, len(combined.columns) - 1, 20)
            ws_all.freeze_panes(1, 0)

    print(f"\nWrote: {out_path}")
    print("Share the 'ALL_combined' sheet (or per-company sheets) with the AI for review.")
    print("\nColour coding:")
    print("  (no colour) = top_amount  — highest MOP, should definitely be tagged correctly")
    print("  yellow      = H_OTHER     — LLM fell back to '其他', probably wrong")
    print("  red         = low_conf    — LLM was unsure (<0.7), likely wrong")


def main():
    parser = argparse.ArgumentParser(description=__doc__, add_help=False)
    parser.add_argument("--entity", default=None)
    args, _ = parser.parse_known_args()
    run(args.entity)


if __name__ == "__main__":
    main()
