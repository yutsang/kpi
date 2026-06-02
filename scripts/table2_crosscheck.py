"""Table 2 cross-check — sanity check 表2 vs kpi_report.xlsx.

表2 (`data/table_2/_extracted.xlsx`) 係項目組嘅 golden source.
呢個 script 比較:

  [A] Per-company total 投資額 (萬澳門元 × 10K → MOP)  vs  kpi_report 25 pivot total
  [B] Per-company × category (Gaming / Non-gaming) breakdown
  [C] Project ID overlap: 表2 projects vs <company>_unique_projects.xlsx
  [D] NG section code distribution (B1.1 / B6 / A1 / ...) — for review
  [E] Top deviation projects (large achievement gaps OR large amount)

Prerequisites:
    kedro run --pipeline=table2_extract      # generates _extracted.xlsx

Run (Windows):
    python scripts/table2_crosscheck.py > table2_crosscheck_output.txt 2>&1
"""
from __future__ import annotations

import sys
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from pathlib import Path
import re
import pandas as pd
import openpyxl

ROOT = Path(".")
T2_EXTRACTED = ROOT / "data" / "table_2" / "_extracted.xlsx"

# 表2 company name → our entity alias
COMPANY_MAP = {
    "Galaxy":    ("company_1", "galaxy"),
    "SJM":       ("company_2", "sjm"),
    "Wynn":      ("company_3", "wynn"),
    "VML":       ("company_4", "vml"),
    "Venetian":  ("company_4", "vml"),
    "Melco":     ("company_5", "melco"),
    "MGM":       ("company_6", "mgm"),
}

# Raw parquet column that directly carries the 表2 project_id value.
# Discovered via scripts/inspect_raw_pretags.py — each entity's source file
# already has a column populated by the project team that aligns with 表2's
# project_id format (B1.15, 項目27, OP002, …). Used by check_c.
RAW_PROJECT_COLS = {
    "Galaxy":   "DICJ Code",                          # e.g. "B1.15", "A1.3"
    "SJM":      "項目",                                # e.g. "項目27"
    "Wynn":     "For计数-Name of Investment Project",  # e.g. "OP002", "CG004"
    "VML":      "項目編號（for 2025執行報告）",         # e.g. "項目38"
    "Venetian": "項目編號（for 2025執行報告）",         # alias of VML
    "Melco":    "承批公司項目序號",                     # e.g. "項目32"
    "MGM":      "project_id",                          # only 11.4% fill — partial match expected
}


def hr(s: str):
    print(f"\n{'='*80}\n  {s}\n{'='*80}")


def safe(s, n: int = 50) -> str:
    if pd.isna(s):
        return ""
    return str(s).strip()[:n]


# ── Load 表2 ────────────────────────────────────────────────────────────────

def load_table2() -> pd.DataFrame | None:
    if not T2_EXTRACTED.exists():
        print(f"❌ 表2 not yet extracted. Run first:")
        print(f"   kedro run --pipeline=table2_extract")
        print(f"   (expects {T2_EXTRACTED.relative_to(ROOT)})")
        return None
    df = pd.read_excel(T2_EXTRACTED, sheet_name="projects")
    # Numeric coercion (in 萬澳門元 — 10K MOP units)
    for c in ["plan_amt_total", "plan_amt_capex", "plan_amt_opex",
              "actual_amt_total", "actual_amt_capex", "actual_amt_opex"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    return df


# ── [A] Per-company total tie-back ──────────────────────────────────────────

def check_a_company_totals(t2: pd.DataFrame):
    hr("[A] Per-company total — 表2 actual_amt_total vs kpi_report 25 pivot")
    print(f"\n  {'Company':<10}  {'表2 actual (萬MOP)':>18}  {'表2 in MOP':>18}  "
          f"{'kpi_report 25 (MOP)':>20}  {'diff (MOP)':>16}  {'diff %':>8}")
    print(f"  {'-'*10}  {'-'*18}  {'-'*18}  {'-'*20}  {'-'*16}  {'-'*8}")

    for t2_company, (company, alias) in COMPANY_MAP.items():
        sub = t2[t2["company"] == t2_company]
        if len(sub) == 0:
            continue
        t2_actual_wan = float(sub["actual_amt_total"].sum())
        t2_actual_mop = t2_actual_wan * 10_000

        # Read kpi_report 25 master pivot total
        # Try tagged_rows.parquet — most authoritative
        kpi_total = None
        tagged_path = ROOT / "data" / alias / "interim" / f"{company}_tagged_rows.parquet"
        if tagged_path.exists():
            try:
                tagged = pd.read_parquet(tagged_path)
                # Find amount column: "amount" first, else fall back to per-entity raw name
                amt_col = None
                for c in ["amount", "Reported Amount(MOP)", "Val/COArea Crcy", "MOP Amt",
                          "Amount - Amended", "amount_mop",
                          "Entry Voucher Amount/ Expense Amount "]:
                    if c in tagged.columns:
                        amt_col = c
                        break
                if amt_col and "report_period" in tagged.columns:
                    rp = tagged["report_period"].astype(str).str.strip()
                    # 表2 actual_amt_total uses ONLY current-year "25" bucket,
                    # NOT including SY adjustments (25_24SY / 25_23SY which are
                    # prior-year carryover corrections). Verified by R5 test:
                    # 5/6 entities tie to 0% diff with "25" only; including SY
                    # introduces -5% to -36% spurious gaps.
                    # MGM ~800M gap is a real pipeline issue (task #19), not
                    # an SY-inclusion problem.
                    mask_25 = rp == "25"
                    kpi_total = float(pd.to_numeric(
                        tagged.loc[mask_25, amt_col], errors="coerce").fillna(0).sum())
                elif amt_col:
                    # No report_period column — sum everything (fallback)
                    kpi_total = float(pd.to_numeric(
                        tagged[amt_col], errors="coerce").fillna(0).sum())
            except Exception as e:
                kpi_total = f"err: {e}"

        if isinstance(kpi_total, (int, float)) and kpi_total > 0:
            diff = t2_actual_mop - kpi_total
            pct = 100 * diff / kpi_total if kpi_total else 0
            print(f"  {t2_company:<10}  {t2_actual_wan:>18,.0f}  {t2_actual_mop:>18,.0f}  "
                  f"{kpi_total:>20,.0f}  {diff:>+16,.0f}  {pct:>+7.1f}%")
        else:
            print(f"  {t2_company:<10}  {t2_actual_wan:>18,.0f}  {t2_actual_mop:>18,.0f}  "
                  f"{'(no data)':>20}  {'-':>16}  {'-':>7}")


# ── [B] Per-company × category breakdown ────────────────────────────────────

def check_b_breakdown(t2: pd.DataFrame):
    hr("[B] Per-company × category breakdown — 表2 plan vs actual (萬澳門元)")
    summary = (
        t2.groupby(["company", "category"], dropna=False)
        .agg(projects=("project_id", "count"),
             plan_total=("plan_amt_total", "sum"),
             actual_total=("actual_amt_total", "sum"),
             plan_capex=("plan_amt_capex", "sum"),
             actual_capex=("actual_amt_capex", "sum"),
             plan_opex=("plan_amt_opex", "sum"),
             actual_opex=("actual_amt_opex", "sum"))
        .reset_index()
    )
    summary["achievement"] = (summary["actual_total"] / summary["plan_total"]).round(4)

    print(f"\n  {'Company':<10} {'Category':<12} {'#proj':>5}  "
          f"{'Plan tot':>11} {'Actual tot':>11}  {'Plan capex':>11} {'Act capex':>11}  "
          f"{'Plan opex':>11} {'Act opex':>11}  {'Achv':>6}")
    print(f"  {'-'*10} {'-'*12} {'-'*5}  {'-'*11} {'-'*11}  {'-'*11} {'-'*11}  "
          f"{'-'*11} {'-'*11}  {'-'*6}")
    for _, r in summary.iterrows():
        achv = r["achievement"]
        achv_s = f"{achv*100:.0f}%" if pd.notna(achv) else "—"
        print(f"  {safe(r['company'], 10):<10} {safe(r['category'], 12):<12} "
              f"{int(r['projects']):>5}  "
              f"{r['plan_total']:>11,.0f} {r['actual_total']:>11,.0f}  "
              f"{r['plan_capex']:>11,.0f} {r['actual_capex']:>11,.0f}  "
              f"{r['plan_opex']:>11,.0f} {r['actual_opex']:>11,.0f}  "
              f"{achv_s:>6}")


# ── [C] Project ID overlap ──────────────────────────────────────────────────

_ID_TOKEN_SPLIT = re.compile(r"[\s\-_/\\,;|()\[\]【】（）、，；]+")
_WS = re.compile(r"\s+")
# Project-ID prefix variants — strip these to normalize "項目73" ↔ "73".
# Some entities' raw col contains "項目NN" while 表2 column has just "NN" (or vice versa).
_PROJ_PREFIXES = ("項目", "项目")


def _id_tokens(s: str) -> list[str]:
    """Split a raw cell value into ID-like tokens (EN+CN delimiters)."""
    if pd.isna(s):
        return []
    parts = _ID_TOKEN_SPLIT.split(str(s).strip())
    return [p.strip().lower() for p in parts if p and p.strip()]


def _id_norm_forms(s) -> set[str]:
    """Return all normalized forms of an ID candidate.

    Generates forms to match across:
      - raw text with whitespace ("項目 35" ↔ "項目35")
      - 項目-prefix differences ("項目73" ↔ "73")
    """
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return set()
    raw = str(s).strip().lower()
    if not raw or raw == "nan":
        return set()
    forms = {raw}
    # whitespace-collapsed form
    no_ws = _WS.sub("", raw)
    forms.add(no_ws)
    # strip 項目-prefix (apply to no_ws form so "項目 35" → "項目35" → "35")
    for pfx in _PROJ_PREFIXES:
        if no_ws.startswith(pfx) and len(no_ws) > len(pfx):
            forms.add(no_ws[len(pfx):])
    forms.discard("")
    return forms


def check_c_project_overlap(t2: pd.DataFrame):
    hr("[C] Project ID overlap — 表2 projects vs raw entity column (project-team tagging)")

    for t2_company, (company, alias) in COMPANY_MAP.items():
        sub = t2[t2["company"] == t2_company]
        if len(sub) == 0:
            continue
        t2_pids = set(sub["project_id"].dropna().astype(str).str.strip().tolist())
        t2_pids.discard("")

        raw_col = RAW_PROJECT_COLS.get(t2_company)
        raw_path = ROOT / "data" / alias / "interim" / f"{company}_raw.parquet"
        if not raw_col or not raw_path.exists():
            print(f"\n  {t2_company} ({alias}): no raw column configured or raw.parquet missing")
            continue

        raw = pd.read_parquet(raw_path, columns=None)
        if raw_col not in raw.columns:
            print(f"\n  {t2_company} ({alias}): raw column {raw_col!r} not in raw.parquet "
                  f"({len(raw.columns)} cols available)")
            continue

        # Build normalized-form set from raw col values.
        # For each cell: tokenize on delimiters, then expand each token through
        # _id_norm_forms (handles whitespace + 項目-prefix variants).
        our_norms: set[str] = set()
        for v in raw[raw_col].dropna():
            # token-level expansion (handles multi-id cells)
            for tok in _id_tokens(v):
                our_norms |= _id_norm_forms(tok)
            # also expand the whole cell (handles "項目 35" with internal space)
            our_norms |= _id_norm_forms(v)
        our_norms.discard("none")

        # 表2 pid matches if ANY of its normalized forms intersects our_norms.
        matched_pids = set()
        for pid in t2_pids:
            if _id_norm_forms(pid) & our_norms:
                matched_pids.add(pid)

        unmatched = t2_pids - matched_pids
        pct = 100 * len(matched_pids) / len(t2_pids) if t2_pids else 0
        n_nonnull = int(raw[raw_col].notna().sum())
        fill_pct = 100 * n_nonnull / len(raw) if len(raw) else 0
        print(f"\n  {t2_company} ({alias})  raw_col={raw_col!r}  "
              f"(filled {n_nonnull:,}/{len(raw):,} = {fill_pct:.1f}%)")
        print(f"    {len(t2_pids)} 表2 pids → matched {len(matched_pids)} ({pct:.0f}%), "
              f"unmatched {len(unmatched)}")
        if unmatched:
            unmatched_df = sub[sub["project_id"].astype(str).str.strip().isin(unmatched)]
            unmatched_df = unmatched_df.sort_values("actual_amt_total", ascending=False)
            print(f"    Top 10 unmatched 表2 projects:")
            for _, r in unmatched_df.head(10).iterrows():
                pid = safe(r["project_id"], 20)
                pname = safe(r.get("actual_project_name") or r.get("plan_project_name", ""), 40)
                amt = float(r.get("actual_amt_total", 0))
                print(f"      {pid:<22} {pname:<42} actual={amt:>10,.0f} 萬MOP")


# ── [D] NG section distribution ──────────────────────────────────────────────

def check_d_section_distribution(t2: pd.DataFrame):
    hr("[D] NG section code distribution (per company × section)")

    dist = (
        t2.groupby(["company", "ng_section_code"], dropna=False)
        .agg(projects=("project_id", "count"),
             actual_total=("actual_amt_total", "sum"))
        .reset_index()
    )

    for company in dist["company"].dropna().unique():
        sub = dist[dist["company"] == company].sort_values("actual_total", ascending=False)
        if len(sub) == 0:
            continue
        print(f"\n  {company} — {len(sub)} sections")
        for _, r in sub.iterrows():
            code = safe(r["ng_section_code"], 12)
            print(f"    {code:<14} projects={int(r['projects']):>3}  "
                  f"actual={r['actual_total']:>10,.0f} 萬MOP")


# ── [E] Large deviation projects ─────────────────────────────────────────────

def check_e_deviations(t2: pd.DataFrame):
    hr("[E] Top deviation projects (large achievement gap OR large amount)")

    # Gaps: achievement < 0.5 OR > 1.5
    t2["_achv"] = pd.to_numeric(
        t2.get("achievement_amt_total", pd.Series([None] * len(t2))), errors="coerce")
    t2["_abs_actual"] = t2["actual_amt_total"].abs()

    underrun = t2[(t2["_achv"].notna()) & (t2["_achv"] < 0.5) & (t2["_abs_actual"] >= 50)]
    underrun = underrun.sort_values("_abs_actual", ascending=False).head(15)
    print(f"\n  UNDERRUN (achievement < 50% AND actual ≥ 50 萬MOP): {len(underrun)} shown")
    for _, r in underrun.iterrows():
        co = safe(r["company"], 10)
        cat = safe(r["category"], 10)
        pid = safe(r["project_id"], 20)
        pname = safe(r.get("actual_project_name") or r.get("plan_project_name", ""), 40)
        achv = r["_achv"] * 100 if pd.notna(r["_achv"]) else 0
        plan = float(r.get("plan_amt_total", 0))
        act = float(r.get("actual_amt_total", 0))
        print(f"    {co:<10} {cat:<10} {pid:<22} {pname:<42}  plan={plan:>8,.0f}  act={act:>8,.0f}  "
              f"achv={achv:>5.0f}%")

    overrun = t2[(t2["_achv"].notna()) & (t2["_achv"] > 1.5) & (t2["_abs_actual"] >= 50)]
    overrun = overrun.sort_values("_abs_actual", ascending=False).head(15)
    print(f"\n  OVERRUN (achievement > 150% AND actual ≥ 50 萬MOP): {len(overrun)} shown")
    for _, r in overrun.iterrows():
        co = safe(r["company"], 10)
        cat = safe(r["category"], 10)
        pid = safe(r["project_id"], 20)
        pname = safe(r.get("actual_project_name") or r.get("plan_project_name", ""), 40)
        achv = r["_achv"] * 100 if pd.notna(r["_achv"]) else 0
        plan = float(r.get("plan_amt_total", 0))
        act = float(r.get("actual_amt_total", 0))
        print(f"    {co:<10} {cat:<10} {pid:<22} {pname:<42}  plan={plan:>8,.0f}  act={act:>8,.0f}  "
              f"achv={achv:>5.0f}%")


def main():
    print("表2 sanity check + cross-check vs project team golden source")
    t2 = load_table2()
    if t2 is None:
        sys.exit(1)
    print(f"\nLoaded 表2: {len(t2)} project rows across {t2['company'].nunique()} companies, "
          f"{t2['category'].nunique()} categories")

    check_a_company_totals(t2)
    check_b_breakdown(t2)
    check_c_project_overlap(t2)
    check_d_section_distribution(t2)
    check_e_deviations(t2)


if __name__ == "__main__":
    main()
