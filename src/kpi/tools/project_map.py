"""Consolidated project → vertical mapping audit for all companies.

Usage:
  python -m kpi.tools.project_map              # flagged rows only (console + Excel)
  python -m kpi.tools.project_map --full       # full project list per company
  python -m kpi.tools.project_map --entity mgm # single company

Output:
  data/audit/project_map_<date>.xlsx  — always written, no truncation issues
  Console: summary + flagged rows only (safe to paste)

Flags:
  LOW_CONF     — confidence < 0.7 (LLM was rejected by NG constraint → forced V_OTHER/0.30)
  V_OTHER_OK   — V_OTHER at confidence ≥ 0.7 (LLM actively chose 其他 — may be correct or wrong)
  NO_VERTICAL  — no vertical assigned at all
  ZERO_AMT     — project has zero total_amount (planned but no actuals, or full offset)
  NEG_AMT      — project has negative total_amount (net reversal)
"""
from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[3]
AUDIT_DIR = ROOT / "data" / "audit"


def _load_one(key: str, alias: str) -> pd.DataFrame | None:
    interim = ROOT / "data" / alias / "interim"
    xlsx = interim / f"{key}_unique_projects.xlsx"
    if not xlsx.exists():
        print(f"  [skip] {key} ({alias}) — no unique_projects.xlsx")
        return None

    df = pd.read_excel(xlsx)
    proj_col = next(
        (c for c in df.columns if c.lower() in
         ("project", "project name", "name of investment project", "項目名稱", "project_name")),
        df.select_dtypes("object").columns[0] if len(df.select_dtypes("object").columns) else df.columns[0],
    )

    out = pd.DataFrame({
        "company": alias,
        "entity_key": key,
        "project": df[proj_col].astype(str).str.strip(),
        "total_amount": pd.to_numeric(
            df.get("total_amount", pd.Series(0, index=df.index)), errors="coerce"
        ).fillna(0),
        "row_count": df.get("row_count", pd.Series(0, index=df.index)),
        "primary_ng": df.get("primary_ng", pd.Series("", index=df.index)).fillna("").astype(str),
        "llm_vertical": df.get("llm_vertical", pd.Series("", index=df.index)).fillna("").astype(str),
        "llm_vertical_label": df.get("llm_vertical_label", pd.Series("", index=df.index)).fillna("").astype(str),
        "llm_confidence": pd.to_numeric(
            df.get("llm_confidence", pd.Series(pd.NA, index=df.index)), errors="coerce"
        ),
        "manual_vertical": df.get("manual_vertical", pd.Series("", index=df.index)).fillna("").astype(str),
        "llm_reasoning": df.get("llm_reasoning", pd.Series("", index=df.index)).fillna("").astype(str),
        "llm_invalid_choice": df.get("llm_invalid_choice", pd.Series("", index=df.index)).fillna("").astype(str),
    })

    out["effective_vertical"] = out.apply(
        lambda r: r["manual_vertical"].strip() if r["manual_vertical"].strip() else r["llm_vertical"], axis=1
    )

    def _flag(r):
        flags = []
        ev = r["effective_vertical"]
        conf = r["llm_confidence"]
        if not ev or ev in ("nan", ""):
            flags.append("NO_VERTICAL")
        elif ev == "(未分類)":
            flags.append("NO_VERTICAL")
        elif pd.notna(conf) and conf < 0.7 and not r["manual_vertical"].strip():
            flags.append("LOW_CONF")
        elif ev == "V_OTHER" and (pd.isna(conf) or conf >= 0.7) and not r["manual_vertical"].strip():
            flags.append("V_OTHER_OK")
        amt = r["total_amount"]
        if amt == 0:
            flags.append("ZERO_AMT")
        elif amt < 0:
            flags.append("NEG_AMT")
        return "|".join(flags) if flags else ""

    out["flag"] = out.apply(_flag, axis=1)
    return out.sort_values("total_amount", ascending=False)


def run(entity_or_alias: str | None = None, full: bool = False):
    from kpi.lib.conf import load_master, load_config

    AUDIT_DIR.mkdir(parents=True, exist_ok=True)

    if entity_or_alias:
        cfg = load_config(entity_or_alias)
        key = cfg["company"]["code"]
        alias = cfg["company"].get("alias") or key
        parts = [_load_one(key, alias)]
    else:
        master = load_master()
        parts = []
        for key, ent in (master.get("companies") or {}).items():
            if not ent or not ent.get("name"):
                continue
            alias = (ent.get("alias") or key).strip()
            parts.append(_load_one(key, alias))

    dfs = [p for p in parts if p is not None]
    if not dfs:
        print("No project data found.")
        return

    combined = pd.concat(dfs, ignore_index=True)

    # ── Console: summary ─────────────────────────────────────────────────────
    print("\n" + "=" * 90)
    print(f"PROJECT → VERTICAL MAPPING  ({len(combined)} projects, {combined['company'].nunique()} companies)")
    print("=" * 90)
    print(f"\n{'Company':<12} {'Projects':>8} {'LOW_CONF':>9} {'V_OTHER_OK':>11} {'ZERO_AMT':>9} {'NEG_AMT':>8} {'Manual':>7}  {'Total MOP':>18}")
    print("-" * 90)
    for company, grp in combined.groupby("company", sort=False):
        n = len(grp)
        n_low = grp["flag"].str.contains("LOW_CONF").sum()
        n_vok = grp["flag"].str.contains("V_OTHER_OK").sum()
        n_zero = grp["flag"].str.contains("ZERO_AMT").sum()
        n_neg = grp["flag"].str.contains("NEG_AMT").sum()
        n_manual = (grp["manual_vertical"].str.strip() != "").sum()
        total = grp["total_amount"].sum()
        print(f"  {company:<10} {n:>8} {n_low:>9} {n_vok:>11} {n_zero:>9} {n_neg:>8} {n_manual:>7}  {total:>18,.0f}")

    # ── Console: flagged rows (always shown — safe to paste) ──────────────────
    flagged = combined[combined["flag"] != ""].copy()
    display_cols = ["company", "project", "total_amount", "primary_ng",
                    "effective_vertical", "llm_confidence", "llm_invalid_choice", "manual_vertical", "flag"]

    def _fmt(df):
        d = df[display_cols].copy()
        d["total_amount"] = d["total_amount"].map(lambda x: f"{x:,.0f}")
        d["llm_confidence"] = d["llm_confidence"].map(lambda x: f"{x:.2f}" if pd.notna(x) else "")
        d["project"] = d["project"].str[:70]
        return d

    if len(flagged):
        print(f"\n── FLAGGED ({len(flagged)} rows) ── LOW_CONF = LLM rejected by NG constraint | V_OTHER_OK = LLM chose 其他 ──")
        pd.set_option("display.max_colwidth", 72)
        pd.set_option("display.width", 240)
        pd.set_option("display.max_rows", 9999)
        for flag_type in ["LOW_CONF", "V_OTHER_OK", "ZERO_AMT", "NEG_AMT", "NO_VERTICAL"]:
            sub = flagged[flagged["flag"].str.contains(flag_type, na=False)]
            if len(sub):
                print(f"\n  [{flag_type}]  {len(sub)} rows")
                print(_fmt(sub).to_string(index=False))
    else:
        print("\n  No flagged projects.")

    if full:
        print(f"\n── FULL LIST ──────────────────────────────────────────────────────────────────────")
        for company, grp in combined.groupby("company", sort=False):
            print(f"\n  ── {company.upper()} ──")
            print(_fmt(grp).to_string(index=False))

    # ── Excel output (always, no truncation) ─────────────────────────────────
    out_path = AUDIT_DIR / f"project_map_{date.today().isoformat()}.xlsx"
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        wb = writer.book
        fmt_low = wb.add_format({"bg_color": "#FFD7D7"})
        fmt_vok = wb.add_format({"bg_color": "#FFF2CC"})
        fmt_zero = wb.add_format({"bg_color": "#E2EFDA"})
        fmt_neg = wb.add_format({"bg_color": "#D9E1F2"})
        fmt_hdr = wb.add_format({"bold": True, "bg_color": "#2F5496", "font_color": "white"})

        excel_cols = ["company", "project", "total_amount", "row_count", "primary_ng",
                      "effective_vertical", "llm_vertical_label", "llm_confidence",
                      "llm_invalid_choice", "llm_reasoning", "manual_vertical", "flag"]

        # Per-company sheets
        for company, grp in combined.groupby("company", sort=False):
            sheet = company[:31]
            grp[excel_cols].to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.sheets[sheet]
            ws.set_column(0, len(excel_cols) - 1, 22)
            ws.set_column(1, 1, 55)   # project name wider
            ws.set_column(9, 9, 40)   # reasoning wider
            ws.freeze_panes(1, 0)
            for row_idx, (_, r) in enumerate(grp[excel_cols].iterrows(), start=1):
                flag = str(r.get("flag", ""))
                if "LOW_CONF" in flag:
                    ws.set_row(row_idx, None, fmt_low)
                elif "V_OTHER_OK" in flag:
                    ws.set_row(row_idx, None, fmt_vok)
                elif "NEG_AMT" in flag:
                    ws.set_row(row_idx, None, fmt_neg)
                elif "ZERO_AMT" in flag:
                    ws.set_row(row_idx, None, fmt_zero)

        # Flagged-only sheet
        if len(flagged):
            flagged[excel_cols].to_excel(writer, sheet_name="FLAGGED", index=False)
            ws_f = writer.sheets["FLAGGED"]
            ws_f.set_column(0, len(excel_cols) - 1, 22)
            ws_f.set_column(1, 1, 55)
            ws_f.set_column(9, 9, 40)
            ws_f.freeze_panes(1, 0)

        # Combined sheet
        combined[excel_cols].to_excel(writer, sheet_name="ALL", index=False)
        ws_all = writer.sheets["ALL"]
        ws_all.set_column(0, len(excel_cols) - 1, 22)
        ws_all.set_column(1, 1, 55)
        ws_all.freeze_panes(1, 0)

    print(f"\nWrote: {out_path}")
    print("Open FLAGGED sheet for review. Colour: red=LOW_CONF, yellow=V_OTHER_OK, green=ZERO_AMT, blue=NEG_AMT")


def main():
    parser = argparse.ArgumentParser(description=__doc__, add_help=False)
    parser.add_argument("--entity", default=None)
    parser.add_argument("--full", action="store_true", help="Also print full project list")
    args, _ = parser.parse_known_args()
    run(args.entity, full=args.full)


if __name__ == "__main__":
    main()
