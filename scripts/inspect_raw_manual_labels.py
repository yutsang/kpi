"""Scan raw parquets for project-team manual labels (ground truth).

User priority: project team often manually labels rows with vertical/horizontal/
capex-opex hints. These are CUSTOMER-PROVIDED → most accurate. We should
detect + apply them BEFORE LLM classification.

For each entity × raw parquet:
  1. List columns containing keywords: 'manual', 'tag', 'category', 'vertical',
     'horizontal', 'label', '分類', '類型', 'comp', 'kp識別'
  2. For each candidate column:
     - Show row coverage (non-empty %)
     - Sample top values + their frequency
     - Detect V_XXX / H_XXX pattern values
     - Show amount impact (sum amount where col is non-empty)

Output: data/_manual_labels_inventory.txt + per-entity console summary

Run:
    python scripts/inspect_raw_manual_labels.py --all
    python scripts/inspect_raw_manual_labels.py --entity sjm
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import pandas as pd
import yaml


ENTITIES = {
    "galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
    "vml":    "company_4", "melco":  "company_5", "mgm": "company_6",
}

LABEL_KEYWORDS = [
    # English
    "manual", "tag", "category", "vertical", "horizontal", "label",
    # Chinese
    "分類", "分类", "類型", "类型", "comp", "kp識別", "kp识别", "識別", "识别",
    "性質", "性质", "scope",
    # KPI-specific
    "ng", "annex",
]


def is_emptyish(s):
    return s.isna() | s.astype(str).str.strip().isin(["", "nan", "<NA>", "-", "0", "None", "NA"])


def analyze_entity(ent: str, com: str) -> list[str]:
    out: list[str] = []
    parquet = Path(f"data/{ent}/interim/{com}_raw.parquet")
    if not parquet.exists():
        return [f"[{ent}] no raw parquet"]

    cfg = yaml.safe_load(Path(f"conf/{com}/parameters.yml").read_text(encoding="utf-8"))
    cols_cfg = cfg.get("columns", {})
    amt_col = cols_cfg.get("amount", "")

    df = pd.read_parquet(parquet)
    total_rows = len(df)
    total_amt = pd.to_numeric(df[amt_col], errors="coerce").fillna(0).sum() if amt_col in df.columns else 0

    out.append("\n" + "=" * 90)
    out.append(f"[{ent}]  rows={total_rows:,}  cols={len(df.columns)}  amount_total={total_amt:,.0f}")
    out.append("=" * 90)

    # Categorize all columns
    out.append(f"\nALL COLUMNS ({len(df.columns)}):")
    for c in df.columns:
        out.append(f"  - {c}")

    # Find label-candidate columns
    candidates = []
    for c in df.columns:
        c_lower = str(c).lower()
        if any(kw.lower() in c_lower for kw in LABEL_KEYWORDS):
            candidates.append(c)

    out.append(f"\nLABEL CANDIDATES ({len(candidates)} cols matched keywords):")
    for c in candidates:
        non_empty_mask = ~is_emptyish(df[c])
        n_non_empty = int(non_empty_mask.sum())
        pct = 100 * n_non_empty / total_rows if total_rows else 0
        amt_impact = pd.to_numeric(df.loc[non_empty_mask, amt_col], errors="coerce").fillna(0).sum() if amt_col in df.columns else 0

        out.append(f"\n  COL: '{c}'  coverage={n_non_empty:,}/{total_rows:,} ({pct:.1f}%)  amount_impact={amt_impact:,.0f}")

        if n_non_empty == 0:
            continue

        # Top values
        vc = df.loc[non_empty_mask, c].astype(str).value_counts().head(15)
        out.append(f"    Top 15 values:")
        for val, cnt in vc.items():
            out.append(f"      {str(val)[:80]:<80}  ({cnt:,} rows)")

        # Check for V_/H_ pattern
        v_pattern = vc.index.astype(str).str.startswith(("V_", "v_")).any()
        h_pattern = vc.index.astype(str).str.startswith(("H_", "h_")).any()
        if v_pattern:
            out.append(f"    *** CONTAINS V_XXX PATTERN — likely project-team vertical label ***")
        if h_pattern:
            out.append(f"    *** CONTAINS H_XXX PATTERN — likely project-team horizontal label ***")

    # Also check OTHER (non-keyword) columns for V_ / H_ patterns
    out.append(f"\nNON-KEYWORD COLUMNS CONTAINING V_/H_ patterns:")
    for c in df.columns:
        if c in candidates:
            continue
        try:
            sample = df[c].dropna().astype(str).head(200)
            if sample.str.startswith(("V_", "H_")).any():
                non_empty_mask = ~is_emptyish(df[c])
                n_non_empty = int(non_empty_mask.sum())
                vc = df.loc[non_empty_mask, c].astype(str).value_counts().head(10)
                v_h_vals = [v for v in vc.index if str(v).startswith(("V_", "H_"))]
                if v_h_vals:
                    out.append(f"  COL: '{c}'  coverage={n_non_empty:,} ({100*n_non_empty/total_rows:.1f}%)")
                    out.append(f"    V/H values: {v_h_vals[:5]}")
        except Exception:
            continue

    return out


def main():
    parser = argparse.ArgumentParser()
    g = parser.add_mutually_exclusive_group(required=True)
    g.add_argument("--entity", choices=list(ENTITIES))
    g.add_argument("--all", action="store_true")
    args = parser.parse_args()

    targets = list(ENTITIES.items()) if args.all else [(args.entity, ENTITIES[args.entity])]
    all_lines = []
    for ent, com in targets:
        print(f">>> {ent}...", flush=True)
        lines = analyze_entity(ent, com)
        all_lines.extend(lines)
        for line in lines:
            print(line, flush=True)
        all_lines.append("")

    out_path = Path("manual_labels_inventory.txt")
    out_path.write_text("\n".join(all_lines), encoding="utf-8")
    print(f"\nFull report: {out_path}")


if __name__ == "__main__":
    main()
