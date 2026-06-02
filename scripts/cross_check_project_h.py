"""Project-level cross-check: project-team signals vs our LLM/rules H tagging.

For each project (e.g. 'Premium Slots A036', 'Eason Chan Concert'), compute:
  1. OUR view: H distribution per LLM/rules (top 3 H by amount)
  2. PROJECT-TEAM signals: NG, capex, NG11 System Code, project name keywords
  3. EXPECTED H: heuristic from above signals
  4. AGREEMENT: does our top-H match expected?

Output (compact, paste-back friendly):

  ## Section A: project-team label coverage
     - List of cols with manual labels + coverage %
     - For Galaxy 25: likely empty (一級/二級 etc not filled), fall back to NG+project name
     - For Wynn 25: comp费用大类 + Annex 2 + Nature of Expenses filled

  ## Section B: per-project H breakdown (top 50 by amount)
     project_name | NG | rows | amt | C/O | LLM top H | LLM split | expected H | match?

  ## Section C: top disagreements
     project | LLM_top | expected | amt-diff | sample_rows

  ## Section D: rule coverage report
     - How many rows did each predominant_rule fire on?
     - How many rows still went to LLM (no rule match)?
     - Top acct_desc that hit LLM (potential new rule candidates)

  ## Section E: H_OTHER projects (high-priority refinement)
     - projects with >30% of amount tagged H_OTHER
     - top sigs in those projects

Run:
  python scripts/cross_check_project_h.py --entity galaxy --year 25
"""
from __future__ import annotations

import argparse
import sys
from collections import Counter, defaultdict
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

EMPTY_VALUES = {"", "nan", "<na>", "-", "0", "none", "na", "null", "n/a"}

# Project name keyword → expected H (for fallback heuristic when project team labels are sparse)
PROJ_NAME_H = [
    ("comp room", "H_HOTEL_ROOM"), ("hotel room", "H_HOTEL_ROOM"),
    ("café", "H_FNB"), ("cafe", "H_FNB"), ("restaurant", "H_FNB"),
    ("bar & lounge", "H_FNB"), ("hot pot", "H_FNB"), ("food festival", "H_FNB"),
    ("food", "H_FNB"), ("gastronomy", "H_FNB"), ("dining", "H_FNB"),
    ("concert", "H_PERFORMER"), ("演唱會", "H_PERFORMER"), ("performer", "H_PERFORMER"),
    ("show", "H_VENUE"),
    ("arena", "H_VENUE"), ("venue", "H_VENUE"), ("convention", "H_VENUE"),
    ("conference", "H_VENUE"), ("event", "H_VENUE"),
    ("sponsorship", "H_ADVERTISING"), ("sport", "H_ADVERTISING"),
    ("marketing", "H_ADVERTISING"), ("promotion", "H_ADVERTISING"),
    ("digital", "H_ADVERTISING"),
    ("gallery", "H_VENUE"), ("museum", "H_VENUE"), ("exhibition", "H_VENUE"),
    ("art space", "H_VENUE"),
    ("renovation", "H_CONSTRUCTION"), ("refurbishment", "H_CONSTRUCTION"),
    ("upgrade", "H_CONSTRUCTION"), ("enhancement", "H_CONSTRUCTION"),
    ("construction", "H_CONSTRUCTION"),
    ("slot", "H_EQUIP"), ("table", "H_EQUIP"), ("gaming equipment", "H_EQUIP"),
    ("hardware", "H_EQUIP"), ("system", "H_EQUIP"), ("technology", "H_EQUIP"),
    ("wifi", "H_EQUIP"), ("smart", "H_EQUIP"),
    ("limo", "H_COMP_OTHER"), ("transport", "H_COMP_OTHER"),
    ("training", "H_LABOR"), ("payroll", "H_LABOR"), ("staff", "H_LABOR"),
    ("office", "H_LABOR"),
    ("wellness", "H_VENUE"), ("spa", "H_COMP_OTHER"),
    ("amusement", "H_VENUE"), ("theme park", "H_VENUE"),
    ("kidz", "H_VENUE"),
    ("recognition", "H_EQUIP"),  # facial recognition system
    ("led", "H_EQUIP"),
]

# NG default H (if no other signal)
NG_DEFAULT_H = {
    "NG0": "H_CONSTRUCTION",
    "NG1": "H_ADVERTISING",
    "NG2": "H_VENUE",
    "NG3": "H_VENUE",  # concert venue rental is largest in Galaxy 25
    "NG4": "H_ADVERTISING",  # sport sponsorship
    "NG5": "H_VENUE",
    "NG6": "H_VENUE",
    "NG7": "H_CONSTRUCTION",
    "NG8": "H_FNB",
    "NG9": "H_ADVERTISING",
    "NG10": "H_CONSTRUCTION",
    "NG11": "H_CONSTRUCTION",
}


def fmt_amt(x: float) -> str:
    a = abs(x)
    sign = "-" if x < 0 else ""
    if a >= 1e9:
        return f"{sign}{a/1e9:.2f}B"
    if a >= 1e6:
        return f"{sign}{a/1e6:.1f}M"
    if a >= 1e3:
        return f"{sign}{a/1e3:.0f}K"
    return f"{sign}{a:.0f}"


def is_empty(s):
    return pd.isna(s) or str(s).strip().lower() in EMPTY_VALUES


def is_emptyish_series(s: pd.Series) -> pd.Series:
    return s.isna() | s.astype(str).str.strip().str.lower().isin(EMPTY_VALUES)


def project_name_to_h(name: str) -> str | None:
    name_l = str(name).lower()
    for kw, h in PROJ_NAME_H:
        if kw in name_l:
            return h
    return None


def _filter_year(df: pd.DataFrame, year_col: str, year: str) -> pd.DataFrame:
    s = df[year_col].astype(str)
    mask = s.str.startswith(year) | (s == f"Yr 20{year}") | (s == f"20{year}")
    return df[mask].copy()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--entity", choices=list(ENTITIES), required=True)
    parser.add_argument("--year", default="25")
    parser.add_argument("--top-projects", type=int, default=50,
                        help="how many top-amount projects to show in Section B")
    parser.add_argument("--top-disagreements", type=int, default=30)
    args = parser.parse_args()

    com = ENTITIES[args.entity]
    parquet = Path(f"data/{args.entity}/output/{com}_kpi_report.parquet")
    if not parquet.exists():
        print(f"❌ {parquet} not found")
        return

    cfg = yaml.safe_load(Path(f"conf/{com}/parameters.yml").read_text(encoding="utf-8"))
    cols_cfg = cfg.get("columns", {})
    amt_col = cols_cfg.get("amount", "")
    ac_col = cols_cfg.get("account_code", "")
    ad_col = cols_cfg.get("account_desc", "")
    dn_col = cols_cfg.get("description", "")
    proj_col = cols_cfg.get("project", "")

    df = pd.read_parquet(parquet)
    year_col = next((c for c in ("report_period", "report_year", "Yr related", "years")
                     if c in df.columns), None)
    if year_col:
        df = _filter_year(df, year_col, args.year)

    ng_col = next((c for c in ("NG11 Category", "NG11 category", "ng11_category")
                   if c in df.columns), None)
    capex_col = next((c for c in ("final_capex_opex", "Capex/Opex", "CAPEX/OPEX",
                                   "Capex / Opex")
                      if c in df.columns), None)

    # Amount col fuzzy fallback — handle trailing whitespace stripped by step5
    if amt_col not in df.columns:
        for c in df.columns:
            if c.strip() == amt_col.strip():
                amt_col = c
                break
    if "horizontal_id" not in df.columns or amt_col not in df.columns:
        print(f"❌ kpi_report parquet missing horizontal_id or amount cols.")
        print(f"   Expected amount col: '{amt_col}'  horizontal_id: {'✓' if 'horizontal_id' in df.columns else '✗'}")
        print(f"   Actual parquet has {len(df.columns)} cols, total rows={len(df):,}")
        print(f"   First 30 cols: {list(df.columns)[:30]}")
        if len(df.columns) > 30:
            print(f"   ...remaining {len(df.columns)-30} cols")
        return

    df["_amt"] = pd.to_numeric(df[amt_col], errors="coerce").fillna(0)
    df["_amt_abs"] = df["_amt"].abs()
    total_rows = len(df)
    total_amt = float(df["_amt"].sum())

    out: list[str] = []
    out.append("=" * 110)
    out.append(f"CROSS-CHECK PROJECT vs RULES — {args.entity}-{args.year}")
    out.append(f"  total_rows={total_rows:,}  total_amount={total_amt:,.0f}")
    out.append("=" * 110)

    # ── Section A: project-team label coverage ──
    out.append("\n## A. PROJECT-TEAM LABEL COLUMN COVERAGE ##")
    label_cols_to_check = [
        # Galaxy 24-only cols (likely empty in 25)
        "一級標簽", "二級標簽", "调整", "Source", "公司名", "公司分类", "性質重分類",
        # Wynn cols
        "comp费用大类", "Breakdown on Comp Expenses in Kind", "Annex 2 Summary Cateogry",
        "Annex 2 Summary Category", "Nature of Expenses", "潜在调整事项-for database",
        "Category", "Annex 15 Foodie Paradise (美食之都) subproject", "Annex 16",
        "事項備註", "其他備註",
        # VML cols
        "會計科目分類", "進一步分類", "Comp類型", "comp支出類型", "分類1.1",
        "Payroll", "Team", "Comp支出",
        # Melco cols
        "支出性質", "支出性质-mapping", "Comp性質-CN", "Comp性質分類",
        "Comp性質分類-EN", "KP識別人工", "spend_category", "JL source",
        # Common
        "NG11 System Code", "DICJ Code",
    ]
    found_cols: list[str] = []
    for c in label_cols_to_check:
        if c in df.columns:
            non_empty = (~is_emptyish_series(df[c])).sum()
            if non_empty == 0:
                out.append(f"  {c:<55} 0% (empty in {args.year})")
            else:
                pct = 100 * non_empty / total_rows
                out.append(f"  {c:<55} {non_empty:>9,} rows ({pct:.1f}%)")
                found_cols.append(c)

    # ── Section B: per-project H breakdown ──
    if not proj_col or proj_col not in df.columns:
        # fallback
        proj_col = next((c for c in ("Project", "Project Name", "project_name",
                                       "Name of Investment Project", "SubProject_Name",
                                       "Sub project")
                          if c in df.columns), None)

    if not proj_col:
        out.append("\n⚠️  no project col — skip per-project analysis")
    else:
        out.append(f"\n## B. PER-PROJECT H BREAKDOWN (top {args.top_projects} by abs amount, project col='{proj_col}') ##")
        # Aggregate
        proj_grp = df.groupby(proj_col)
        proj_rows = []
        for pname, sub in proj_grp:
            pname_s = str(pname) if not pd.isna(pname) else "(NaN)"
            p_amt = float(sub["_amt"].sum())
            p_abs = float(sub["_amt_abs"].sum())
            p_rows = len(sub)
            # NG modal
            ng_modal = ""
            if ng_col:
                ngc = sub[ng_col].astype(str).value_counts()
                ng_modal = ngc.index[0] if len(ngc) else ""
            # Capex/opex
            cap_split = ""
            if capex_col:
                ce_lower = sub[capex_col].astype(str).str.lower()
                cap_amt = float(sub.loc[ce_lower.str.contains("capex", na=False), "_amt"].sum())
                opex_amt = float(sub.loc[ce_lower.str.contains("opex", na=False), "_amt"].sum())
                tot_ce = abs(cap_amt) + abs(opex_amt)
                cap_pct = 100 * abs(cap_amt) / tot_ce if tot_ce else 0
                cap_split = f"C{cap_pct:.0f}%"
            # Top H by amount
            h_grp = sub.groupby("horizontal_id")["_amt"].sum().sort_values(key=lambda s: s.abs(), ascending=False)
            top_h = h_grp.index[0] if len(h_grp) else ""
            top_h_pct = 100 * abs(h_grp.iloc[0]) / abs(p_amt) if p_amt else 0 if len(h_grp) else 0
            # H summary string (top 3)
            h_str_parts = []
            for h, a in list(h_grp.items())[:3]:
                pct = 100 * abs(a) / p_abs if p_abs else 0
                h_short = str(h).replace("H_", "")
                h_str_parts.append(f"{h_short}:{pct:.0f}%")
            h_str = "/".join(h_str_parts)
            # Expected H heuristic
            exp_from_name = project_name_to_h(pname_s)
            exp_from_ng = NG_DEFAULT_H.get(ng_modal, "?")
            expected = exp_from_name or exp_from_ng
            match = "✓" if expected and str(top_h).endswith(expected.replace("H_", "")) else ("?" if not expected else "✗")
            proj_rows.append({
                "project": pname_s,
                "ng": ng_modal,
                "rows": p_rows,
                "amt": p_amt,
                "abs": p_abs,
                "cap": cap_split,
                "top_h": top_h,
                "h_summary": h_str,
                "expected": expected,
                "match": match,
            })
        proj_rows.sort(key=lambda x: -x["abs"])
        top_proj = proj_rows[:args.top_projects]
        out.append(f"  {'project':<60} {'NG':<5} {'rows':>6} {'amt':>9} {'cap':<5} "
                   f"{'top_H':<14} {'breakdown':<32} {'expected':<14} m")
        out.append("  " + "-" * 160)
        for p in top_proj:
            pn = p["project"][:60]
            out.append(f"  {pn:<60} {p['ng']:<5} {p['rows']:>6,} {fmt_amt(p['amt']):>9} "
                       f"{p['cap']:<5} {str(p['top_h']):<14} {p['h_summary']:<32} "
                       f"{str(p['expected'] or '?'):<14} {p['match']}")

        # ── Section C: top disagreements ──
        out.append(f"\n## C. TOP DISAGREEMENTS (LLM_top != expected, projects ranked by abs amt) ##")
        disagr = [p for p in proj_rows if p["match"] == "✗"]
        out.append(f"  total disagreement projects: {len(disagr):,} ({len(disagr)/len(proj_rows)*100 if proj_rows else 0:.1f}% of projects)")
        out.append(f"  {'project':<60} {'NG':<5} {'amt':>9} {'LLM_top':<14} {'expected':<14}")
        out.append("  " + "-" * 110)
        for p in disagr[:args.top_disagreements]:
            pn = p["project"][:60]
            out.append(f"  {pn:<60} {p['ng']:<5} {fmt_amt(p['amt']):>9} "
                       f"{str(p['top_h']):<14} {str(p['expected'] or '?'):<14}")

        # ── Section E: H_OTHER projects ──
        out.append(f"\n## E. PROJECTS WITH >=30% H_OTHER (high-priority refinement) ##")
        other_projs = []
        for pname, sub in proj_grp:
            other_amt = float(sub[sub["horizontal_id"] == "H_OTHER"]["_amt"].sum())
            tot_p = float(sub["_amt"].sum())
            if tot_p and abs(other_amt) / abs(tot_p) >= 0.30:
                other_projs.append({
                    "project": str(pname)[:60],
                    "ng": sub[ng_col].mode().iloc[0] if ng_col else "",
                    "other_amt": other_amt,
                    "tot_amt": tot_p,
                    "pct": 100 * abs(other_amt) / abs(tot_p),
                })
        other_projs.sort(key=lambda x: -abs(x["other_amt"]))
        out.append(f"  {'project':<60} {'NG':<5} {'other_amt':>11} {'tot_amt':>11} {'pct':>6}")
        out.append("  " + "-" * 100)
        for p in other_projs[:20]:
            out.append(f"  {p['project']:<60} {p['ng']:<5} {fmt_amt(p['other_amt']):>11} "
                       f"{fmt_amt(p['tot_amt']):>11} {p['pct']:>5.1f}%")

    # ── Section D: rule coverage report ──
    out.append(f"\n## D. RULE COVERAGE — top UNRULED acct_desc patterns (LLM-only) ##")
    if ad_col in df.columns:
        # Heuristic: acct_desc with high count + LLM-only (we can't directly know which rows hit rules,
        # but we can look for patterns where LLM gave H_OTHER — these missed both rule + LLM)
        ad_grp = (df[df["horizontal_id"] == "H_OTHER"]
                  .groupby(ad_col)["_amt"]
                  .agg(["count", "sum"])
                  .sort_values("sum", key=lambda s: s.abs(), ascending=False)
                  .head(30))
        out.append(f"  Top 30 acct_desc that got H_OTHER (rule + LLM both missed):")
        out.append(f"  {'acct_desc':<60} {'rows':>9} {'amount':>12}")
        out.append("  " + "-" * 90)
        for ad_v, row in ad_grp.iterrows():
            out.append(f"  {str(ad_v)[:60]:<60} {int(row['count']):>9,} {row['sum']:>12,.0f}")

    out_path = Path(f"cross_check_{args.entity}_{args.year}.txt")
    out_path.write_text("\n".join(out), encoding="utf-8")
    for line in out:
        print(line)
    print(f"\n✓ wrote {out_path}  ({out_path.stat().st_size // 1024} KB, {len(out):,} lines)")


if __name__ == "__main__":
    main()
