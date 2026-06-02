"""Deep year-specific profile of raw parquet — understand raw files anew.

For ONE entity × ONE year (24 or 25), produces a comprehensive profile:

  SECTION 1 — SCHEMA + COVERAGE
    Every column, dtype, % non-empty, total amount impact

  SECTION 2 — BASELINES
    Total rows + amount
    Capex/Opex breakdown (every capex_opex-like column found)
    Comp-scope breakdown (rows with comp indicator)
    NG (NG0-NG11) distribution if present
    V (vertical) distribution

  SECTION 3 — V-COLS vs H-COLS classification
    Auto-detect which cols describe vertical (project type/nature)
    vs which cols describe horizontal (cost category)

  SECTION 4 — H-COL value → MAPPING status
    For each H-col candidate, show:
      - Top values (rows + amount + capex/opex split)
      - Whether MAPPED in apply_manual_labels MAPPINGS
      - Unmapped values with HIGH amount → ADD candidates

  SECTION 5 — CAPEX vs OPEX per H-col value
    User insight: capex/opex split varies per label value.
    Helps verify mapping correctness (capex-heavy → likely H_CONSTRUCTION/H_EQUIP)

  SECTION 6 — PROMPT/RULE SUGGESTIONS
    Based on findings, output actionable text to:
    (a) Add to LLM system prompt (high-purity acct/desc patterns per H)
    (b) Add as predominant_rules in conf/<ent>/parameters.yml

Output: deep_profile_{ent}_{year}.txt

Run:
  python scripts/deep_profile_entity_year.py --entity galaxy --year 25
  python scripts/deep_profile_entity_year.py --all-25
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

from apply_manual_labels_as_overrides import MAPPINGS, MULTI_COL_RULES


ENTITIES = {
    "galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
    "vml":    "company_4", "melco":  "company_5", "mgm": "company_6",
}

EMPTY_VALUES = {"", "nan", "<na>", "-", "0", "none", "na", "null", "n/a", "0.0"}

COMP_KEYWORDS = ["comp", "complimentary", "贈", "招待", "免費", "免费",
                 "in kind", "vik"]

V_HINT_KEYWORDS = ["項目", "项目", "性質", "性质", "nature", "category",
                   "initiative", "scope", "業務", "业务"]

H_HINT_KEYWORDS = ["費用", "费用", "expense", "cost", "支出", "payroll",
                   "comp", "rental", "lease", "labor", "salary",
                   "ledger", "spend", "性质-mapping", "ng11"]

CAPEX_OPEX_KEYWORDS = ["capex", "opex", "性質重分類", "性质重分类",
                       "report_capex_opex", "final_capex_opex"]


def is_emptyish_series(s: pd.Series) -> pd.Series:
    return s.isna() | s.astype(str).str.strip().str.lower().isin(EMPTY_VALUES)


def fmt_amt(x: float) -> str:
    if abs(x) >= 1e9:
        return f"{x/1e9:>7.2f}B"
    if abs(x) >= 1e6:
        return f"{x/1e6:>7.2f}M"
    if abs(x) >= 1e3:
        return f"{x/1e3:>7.1f}K"
    return f"{x:>9,.0f}"


def detect_year_col(df: pd.DataFrame) -> str | None:
    for c in ("report_period", "report_year", "Yr related", "years"):
        if c in df.columns:
            return c
    return None


def filter_by_year(df: pd.DataFrame, year_col: str, year: str) -> pd.DataFrame:
    """Filter df to rows matching year (e.g. '25', '24'). Match by leading 2 chars
    OR by Chinese 'Yr 2025' style OR by year integer 2025."""
    s = df[year_col].astype(str)
    mask = (
        s.str.startswith(year)
        | (s == f"Yr 20{year}")
        | (s == f"20{year}")
        | (s == f"20{year}-01-01 00:00:00")  # melco style
        | s.str.contains(f"20{year}年", na=False)
    )
    return df[mask]


def col_categorize(df: pd.DataFrame, col: str) -> str:
    """Return one of: 'amount', 'id', 'timestamp', 'numeric', 'V', 'H', 'capex_opex', 'other'."""
    cl = str(col).lower()
    if any(k in cl for k in ("amount", "amt")):
        return "amount"
    if any(k in cl for k in ("uid", "_id", "row_id", "voucher", "ref ", "ref number")):
        return "id"
    if any(k in cl for k in ("date", "time", "period", "year", "month")):
        return "timestamp"
    if pd.api.types.is_numeric_dtype(df[col]) and not pd.api.types.is_bool_dtype(df[col]):
        return "numeric"
    if any(k.lower() in cl for k in CAPEX_OPEX_KEYWORDS):
        return "capex_opex"
    # V/H hints
    has_v = any(k.lower() in cl for k in V_HINT_KEYWORDS)
    has_h = any(k.lower() in cl for k in H_HINT_KEYWORDS)
    if has_h and not has_v:
        return "H"
    if has_v and not has_h:
        return "V"
    if has_v and has_h:
        return "V/H"
    return "other"


def get_amt_series(df: pd.DataFrame, amt_col: str) -> pd.Series:
    return pd.to_numeric(df[amt_col], errors="coerce").fillna(0)


def profile_entity_year(ent: str, com: str, year: str) -> list[str]:
    out: list[str] = []
    parquet = Path(f"data/{ent}/interim/{com}_raw.parquet")
    if not parquet.exists():
        return [f"\n[{ent}-{year}] no raw parquet — kedro step0_5 not done"]

    cfg = yaml.safe_load(Path(f"conf/{com}/parameters.yml").read_text(encoding="utf-8"))
    cols_cfg = cfg.get("columns", {})
    amt_col = cols_cfg.get("amount", "")

    df_full = pd.read_parquet(parquet)
    year_col = detect_year_col(df_full)
    if not year_col:
        return [f"\n[{ent}-{year}] no year column detected — cols: {list(df_full.columns)[:10]}..."]
    df = filter_by_year(df_full, year_col, year)
    if len(df) == 0:
        return [f"\n[{ent}-{year}] no rows for year {year} (year_col='{year_col}')"]

    amt = get_amt_series(df, amt_col) if amt_col in df.columns else pd.Series([0] * len(df))
    total_rows = len(df)
    total_amt = float(amt.sum())

    out.append("\n" + "=" * 100)
    out.append(f"[{ent}-{year}]  rows={total_rows:,}  total_amount={total_amt:,.0f}  "
               f"({len(df)/len(df_full)*100:.1f}% of full {len(df_full):,})")
    out.append(f"  source: {parquet.name}, year_col='{year_col}', amount_col='{amt_col}'")
    out.append("=" * 100)

    # ─────────────────────────────────────────────────────────────────────────
    # SECTION 1 — SCHEMA + COVERAGE (ALL columns, no filter)
    # ─────────────────────────────────────────────────────────────────────────
    out.append(f"\n### SECTION 1 — SCHEMA + COVERAGE ({len(df.columns)} cols) ###")
    out.append(f"{'col':<55} {'cat':<10} {'cover%':>7} {'uniq':>6} {'amt_impact':>12}")
    out.append("-" * 95)
    col_categories: dict[str, str] = {}
    zero_cov_cols: list[str] = []
    for col in df.columns:
        cat = col_categorize(df, col)
        col_categories[col] = cat
        if cat in ("amount", "numeric"):
            non_empty = ~df[col].isna() if pd.api.types.is_numeric_dtype(df[col]) else (~is_emptyish_series(df[col]))
        else:
            non_empty = ~is_emptyish_series(df[col])
        n_ne = int(non_empty.sum())
        cov_pct = 100 * n_ne / total_rows if total_rows else 0
        try:
            n_uniq = df.loc[non_empty, col].astype(str).nunique() if n_ne else 0
        except Exception:
            n_uniq = 0
        amt_impact = float(amt[non_empty].sum()) if n_ne else 0
        # Show ALL cols regardless of coverage (so user can see hidden manual label cols)
        out.append(f"{str(col)[:55]:<55} {cat:<10} {cov_pct:>6.1f}% {n_uniq:>6,} {fmt_amt(amt_impact):>12}")
        if cov_pct == 0:
            zero_cov_cols.append(str(col))

    # Highlight zero-coverage cols — these might be manual-label cols that project
    # team left empty for this year. Important to surface (cols exist but no data).
    if zero_cov_cols:
        out.append(f"\n  ⚠️  Zero-coverage cols ({len(zero_cov_cols)}): may be manual-label cols project team "
                   f"left empty for this year — check if project team plans to fill them:")
        for c in zero_cov_cols:
            out.append(f"    - {c}")

    # ─────────────────────────────────────────────────────────────────────────
    # SECTION 2 — BASELINES
    # ─────────────────────────────────────────────────────────────────────────
    out.append(f"\n### SECTION 2 — BASELINES ###")

    # 2a. Capex/Opex breakdown
    out.append(f"\n  CAPEX/OPEX (per detected col):")
    capex_cols = [c for c, cat in col_categories.items() if cat == "capex_opex"]
    for c in capex_cols:
        non_empty = ~is_emptyish_series(df[c])
        if non_empty.sum() == 0:
            continue
        breakdown = df.loc[non_empty].groupby(c).apply(
            lambda g: pd.Series({"rows": len(g), "amt": amt.loc[g.index].sum()})
        ).sort_values("amt", ascending=False)
        out.append(f"\n  [{c}]:")
        for v, row in breakdown.iterrows():
            pct = 100 * row["amt"] / total_amt if total_amt else 0
            out.append(f"    {str(v)[:30]:<30}  {int(row['rows']):>7,} rows  "
                       f"{row['amt']:>15,.0f}  ({pct:>5.1f}%)")

    # 2b. Comp baseline (rows with comp indicator)
    out.append(f"\n  COMP BASELINE (multiple detection methods):")
    ac_col = cols_cfg.get("account_desc", "")
    dc_col = cols_cfg.get("description", "")
    comp_methods = []

    # Method A: account_desc contains comp keyword
    if ac_col in df.columns:
        ad_s = df[ac_col].astype(str).str.lower()
        m_ad = pd.Series(False, index=df.index)
        for kw in COMP_KEYWORDS:
            m_ad |= ad_s.str.contains(kw, na=False, regex=False)
        comp_methods.append(("account_desc has comp-keyword", m_ad))

    # Method B: description contains comp keyword
    if dc_col in df.columns:
        dc_s = df[dc_col].astype(str).str.lower()
        m_dc = pd.Series(False, index=df.index)
        for kw in COMP_KEYWORDS:
            m_dc |= dc_s.str.contains(kw, na=False, regex=False)
        comp_methods.append(("description has comp-keyword", m_dc))

    # Method C: per-entity comp-specific columns
    comp_col_candidates = ["Comp支出", "Comp", "comp费用大类", "Breakdown on Comp Expenses in Kind",
                            "Comp類型", "comp支出類型", "Comp性質-CN",
                            "Comp性質-CN（N/A為Net off及不適用），待確認kp識別，客戶未識別部分）",
                            "KP識別Comp"]
    for cc in comp_col_candidates:
        if cc in df.columns:
            m_cc = ~is_emptyish_series(df[cc])
            comp_methods.append((f"col '{cc}' filled", m_cc))

    for label, mask in comp_methods:
        n = int(mask.sum())
        a = float(amt[mask].sum())
        if n:
            out.append(f"    {label:<70}  {n:>7,} rows  {a:>15,.0f}  ({100*a/total_amt if total_amt else 0:>5.2f}%)")

    # 2c. NG distribution (Galaxy / MGM)
    for ng_col in ("NG11 Category", "NG11 category", "ng11_category"):
        if ng_col in df.columns:
            non_empty = ~is_emptyish_series(df[ng_col])
            if non_empty.sum() > 0:
                out.append(f"\n  NG DISTRIBUTION ('{ng_col}'):")
                ng_grp = df.loc[non_empty].groupby(ng_col).apply(
                    lambda g: pd.Series({"rows": len(g), "amt": amt.loc[g.index].sum()})
                ).sort_values("amt", ascending=False).head(15)
                for v, row in ng_grp.iterrows():
                    pct = 100 * row["amt"] / total_amt if total_amt else 0
                    if pct >= 0.1:
                        out.append(f"    {str(v)[:30]:<30}  {int(row['rows']):>7,}  "
                                   f"{row['amt']:>15,.0f}  ({pct:>5.1f}%)")
            break  # only show one

    # 2d. V (vertical/project) distribution
    v_col_candidates = ["項目性質", "nature", "Initiative Name", "initiative ID"]
    for v_col in v_col_candidates:
        if v_col in df.columns:
            non_empty = ~is_emptyish_series(df[v_col])
            if non_empty.sum() > 0:
                out.append(f"\n  V (vertical/project) DISTRIBUTION ('{v_col}'):")
                vg = df.loc[non_empty].groupby(v_col).apply(
                    lambda g: pd.Series({"rows": len(g), "amt": amt.loc[g.index].sum()})
                ).sort_values("amt", ascending=False).head(10)
                for v, row in vg.iterrows():
                    pct = 100 * row["amt"] / total_amt if total_amt else 0
                    if pct >= 0.5:
                        out.append(f"    {str(v)[:40]:<40}  {int(row['rows']):>7,}  "
                                   f"{row['amt']:>15,.0f}  ({pct:>5.1f}%)")
                break

    # ─────────────────────────────────────────────────────────────────────────
    # SECTION 3 — V vs H column classification
    # ─────────────────────────────────────────────────────────────────────────
    out.append(f"\n### SECTION 3 — COLUMN ROLE CLASSIFICATION ###")
    v_cols = [c for c, cat in col_categories.items() if cat == "V"]
    h_cols = [c for c, cat in col_categories.items() if cat == "H"]
    vh_cols = [c for c, cat in col_categories.items() if cat == "V/H"]
    out.append(f"  V-cols (vertical labels — DO NOT auto-map to H): {len(v_cols)}")
    for c in v_cols:
        out.append(f"    - {c}")
    out.append(f"\n  H-cols (horizontal labels — CAN map to H): {len(h_cols)}")
    for c in h_cols:
        out.append(f"    - {c}")
    out.append(f"\n  V/H ambiguous: {len(vh_cols)}")
    for c in vh_cols:
        out.append(f"    - {c}")

    # ─────────────────────────────────────────────────────────────────────────
    # SECTION 4 — H-col value → MAPPING status (CAPEX/OPEX split per value)
    # ─────────────────────────────────────────────────────────────────────────
    out.append(f"\n### SECTION 4 — H-COL VALUE BREAKDOWN ###")
    mapped_cols = MAPPINGS.get(ent, {})
    capex_col = next((c for c, cat in col_categories.items() if cat == "capex_opex"), None)

    # Show all H-cols (already MAPPED ones, to detect coverage gaps)
    for col in h_cols + vh_cols:
        non_empty = ~is_emptyish_series(df[col])
        n_ne = int(non_empty.sum())
        if n_ne == 0:
            continue
        mapped_values = mapped_cols.get(col, {})
        mapped_flag = " [MAPPED]" if mapped_values else " [UNMAPPED]"
        out.append(f"\n  COL: '{col}'  coverage={n_ne:,} ({100*n_ne/total_rows:.1f}%){mapped_flag}")

        vc = df.loc[non_empty].groupby(col).apply(
            lambda g: pd.Series({"rows": len(g), "amt": amt.loc[g.index].sum()})
        ).sort_values("amt", key=lambda s: s.abs(), ascending=False).head(20)

        if capex_col and capex_col in df.columns:
            out.append(f"    {'value':<50} {'rows':>7} {'amount':>15} {'capex%':>7} {'mapped→H':<14}")
        else:
            out.append(f"    {'value':<50} {'rows':>7} {'amount':>15} {'mapped→H':<14}")
        for v, row in vc.iterrows():
            v_str = str(v)[:50]
            mapped_h = mapped_values.get(str(v), "—")
            mark = " ⭐ ADD" if mapped_h == "—" and abs(row["amt"]) > 10_000_000 else ""
            if capex_col and capex_col in df.columns:
                mask_v = (df[col].astype(str) == str(v))
                sub = df.loc[mask_v]
                if len(sub) > 0:
                    capex_str_lower = sub[capex_col].astype(str).str.lower()
                    capex_amt = float(amt.loc[mask_v & capex_str_lower.str.contains("capex", na=False)].sum())
                    opex_amt = float(amt.loc[mask_v & capex_str_lower.str.contains("opex", na=False)].sum())
                    total_ce = capex_amt + opex_amt
                    capex_pct = (100 * capex_amt / total_ce) if total_ce else 0
                else:
                    capex_pct = 0
                out.append(f"    {v_str:<50} {int(row['rows']):>7,} {row['amt']:>15,.0f} "
                           f"{capex_pct:>6.0f}% {mapped_h:<14}{mark}")
            else:
                out.append(f"    {v_str:<50} {int(row['rows']):>7,} {row['amt']:>15,.0f}  "
                           f"{mapped_h:<14}{mark}")

    # ─────────────────────────────────────────────────────────────────────────
    # SECTION 5 — PROMPT / RULE SUGGESTIONS
    # ─────────────────────────────────────────────────────────────────────────
    out.append(f"\n### SECTION 5 — ACTIONABLE SUGGESTIONS ###")

    # 5a. Suggest mapping additions for high-amount unmapped values in MAPPED H-cols
    add_candidates: list[tuple[str, str, str, int, float]] = []  # (col, value, suggested_h, rows, amt)
    for col, value_map in mapped_cols.items():
        if col not in df.columns:
            continue
        non_empty = ~is_emptyish_series(df[col])
        if non_empty.sum() == 0:
            continue
        for v in df.loc[non_empty, col].astype(str).unique():
            if v in value_map:
                continue
            mask = (df[col].astype(str) == v)
            row_amt = float(amt[mask].sum())
            if abs(row_amt) >= 5_000_000:
                # heuristic suggestion from value text
                from inspect_multi_column_labels import suggest_h_from_text  # type: ignore
                suggested = suggest_h_from_text(v) or "?"
                add_candidates.append((col, v, suggested, int(mask.sum()), row_amt))

    if add_candidates:
        out.append(f"\n  (a) MAPPING ADDITIONS — high-$ unmapped values in already-MAPPED H-cols:")
        for col, v, sug, n, a in sorted(add_candidates, key=lambda x: -abs(x[4]))[:20]:
            out.append(f"    MAPPINGS['{ent}']['{col}']['{v}'] = '{sug}'")
            out.append(f"      → impact: {n:,} rows  {a:,.0f} MOP")

    # 5b. Suggest LLM system prompt enhancement (high-purity acct/desc patterns per H)
    out.append(f"\n  (b) LLM PROMPT ENHANCEMENT — acct_desc patterns with >80% manual purity:")
    if ac_col in df.columns:
        # For each currently MAPPED H, find acct keywords highly correlated
        # We use the manual labels themselves as "ground truth" purity
        labeled_h_per_row = pd.Series([None] * len(df), index=df.index, dtype="object")
        for col, value_map in mapped_cols.items():
            if col not in df.columns:
                continue
            for v, h in value_map.items():
                mask = (labeled_h_per_row.isna()) & (df[col].astype(str) == v)
                labeled_h_per_row.loc[mask] = h

        labeled_df = df[labeled_h_per_row.notna()].copy()
        labeled_df["_h"] = labeled_h_per_row[labeled_h_per_row.notna()]
        if len(labeled_df) > 100:
            for h in sorted(labeled_df["_h"].unique()):
                sub = labeled_df[labeled_df["_h"] == h]
                if len(sub) < 5:
                    continue
                # Find acct values with high purity for this H
                vc = sub[ac_col].astype(str).value_counts().head(10)
                patterns = []
                for ac_val, cnt in vc.items():
                    if not ac_val or ac_val.lower() in EMPTY_VALUES:
                        continue
                    # purity: of all labeled rows with this acct, what % are this H
                    all_rows_this_ac = labeled_df[labeled_df[ac_col].astype(str) == ac_val]
                    if len(all_rows_this_ac) < 3:
                        continue
                    purity = (all_rows_this_ac["_h"] == h).mean()
                    if purity >= 0.80:
                        patterns.append((ac_val, cnt, purity))
                if patterns:
                    out.append(f"\n    → {h}:")
                    for ac_val, cnt, pur in patterns[:8]:
                        out.append(f"      acct LIKE '{str(ac_val)[:50]}'  ({cnt:,} rows, {pur*100:.0f}% pure)")

    return out


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--entity", choices=list(ENTITIES))
    parser.add_argument("--year", default="25", help="year suffix to filter, default 25")
    parser.add_argument("--all-25", action="store_true", help="profile all 6 entities for year 25")
    parser.add_argument("--all-24", action="store_true", help="profile all 6 entities for year 24")
    args = parser.parse_args()

    if args.all_25:
        targets = [(e, c, "25") for e, c in ENTITIES.items()]
    elif args.all_24:
        targets = [(e, c, "24") for e, c in ENTITIES.items()]
    elif args.entity:
        targets = [(args.entity, ENTITIES[args.entity], args.year)]
    else:
        parser.error("specify --entity OR --all-25 OR --all-24")

    for ent, com, year in targets:
        print(f"\n>>> profiling {ent}-{year}...", flush=True)
        lines = profile_entity_year(ent, com, year)
        out_path = Path(f"deep_profile_{ent}_{year}.txt")
        out_path.write_text("\n".join(lines), encoding="utf-8")
        for line in lines:
            print(line, flush=True)
        print(f"\n✓ wrote {out_path}")


if __name__ == "__main__":
    main()
