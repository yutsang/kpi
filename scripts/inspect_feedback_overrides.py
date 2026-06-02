"""Inspect feedback.xlsx for problematic H_OTHER overrides that block rules.

After Round 5 + Round 6 conf patches, Galaxy 25 cross-check showed NO change
in H_OTHER pollution (Repair & Maint, Corporate Service, etc still tagged
H_OTHER). Most likely cause: feedback.xlsx has these sigs labeled as H_OTHER
from earlier rounds. Feedback overrides take priority over predominant_rules,
so new rules cannot fire.

This script:
  1. Loads feedback.xlsx
  2. Finds sigs where correct_horizontal == 'H_OTHER' AND the signature contains
     known patterns (Repair, Corporate, Sponsorship Fee, restaurant project, etc.)
  3. Lists them so user can decide to:
     (a) Clear those rows (revert to LLM/rules)
     (b) Manually update correct_horizontal to the right H
     (c) Leave as-is if user explicitly wanted H_OTHER

Run:
  python scripts/inspect_feedback_overrides.py --entity galaxy
  python scripts/inspect_feedback_overrides.py --entity galaxy --auto-fix
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


ENTITIES = {
    "galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
    "vml":    "company_4", "melco":  "company_5", "mgm": "company_6",
}

# Patterns from Round 5/6 Galaxy conf patches — if these are tagged H_OTHER in
# feedback.xlsx, they should be reverted/updated to match new rules.
KNOWN_PATTERNS = {
    # Round 5 patterns
    "Repair & Maint": "H_MAINTENANCE",
    "GENERAL ENG SVC": "H_MAINTENANCE",
    "ELEC SYS SVC": "H_MAINTENANCE",
    "LIGHTING SYS SVC": "H_MAINTENANCE",
    "HVAC SVC": "H_MAINTENANCE",
    "Corporate Service": "H_PROFESSIONAL",
    "Corp.Serv": "H_PROFESSIONAL",
    "Sponsorship Fee": "H_ADVERTISING",
    "Donations": "H_SPONSORSHIP",
    "Donation to": "H_SPONSORSHIP",
    # Round 6 — FA prefix overrides
    "FA - Uniform": "H_EQUIP",
    "FA - Chinaware": "H_EQUIP",
    "FA - Glassware": "H_EQUIP",
    "FA - Linen": "H_EQUIP",
}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--entity", choices=list(ENTITIES), required=True)
    parser.add_argument("--auto-fix", action="store_true",
                        help="rewrite feedback.xlsx with corrected H per KNOWN_PATTERNS")
    parser.add_argument("--clear-only", action="store_true",
                        help="clear matching rows' correct_horizontal (so rules can fire)")
    args = parser.parse_args()

    com = ENTITIES[args.entity]
    fb_path = Path(f"data/{args.entity}/output/{com}_feedback.xlsx")
    if not fb_path.exists():
        print(f"❌ {fb_path} not found")
        return

    fb = pd.read_excel(fb_path)
    print(f"Loaded {fb_path.name}  rows={len(fb):,}  cols={list(fb.columns)}")

    if "signature" not in fb.columns or "correct_horizontal" not in fb.columns:
        print(f"❌ feedback.xlsx missing 'signature' or 'correct_horizontal' col")
        return

    # Show distribution of correct_horizontal values
    h_dist = fb["correct_horizontal"].astype(str).value_counts()
    print(f"\nFeedback correct_horizontal distribution:")
    for h, cnt in h_dist.head(20).items():
        print(f"  {h:<20} {cnt:>6,}")

    # Find H_OTHER overrides that match KNOWN_PATTERNS
    other_rows = fb[fb["correct_horizontal"].astype(str) == "H_OTHER"]
    print(f"\nTotal H_OTHER overrides: {len(other_rows):,}")
    # Show notes distribution (clue: which override source set H_OTHER)
    if "notes" in other_rows.columns:
        nc = other_rows["notes"].astype(str).str[:50].value_counts().head(10)
        print(f"  Top notes for H_OTHER overrides:")
        for note, cnt in nc.items():
            print(f"    [{note[:50]}] {cnt:>5,}")

    problematic = []
    for pattern, expected_h in KNOWN_PATTERNS.items():
        mask = other_rows["signature"].astype(str).str.contains(pattern, case=False, na=False, regex=False)
        matched = other_rows[mask]
        if len(matched):
            for _, r in matched.iterrows():
                problematic.append({
                    "row_idx": r.name,
                    "signature": str(r["signature"])[:80],
                    "current_h": "H_OTHER",
                    "expected_h": expected_h,
                    "pattern": pattern,
                })

    # Amount estimate: look up sigs in kpi_report.parquet (if exists)
    rep_path = Path(f"data/{args.entity}/output/{com}_kpi_report.parquet")
    sig_to_amt: dict[str, float] = {}
    sig_to_rows: dict[str, int] = {}
    if rep_path.exists():
        try:
            rep = pd.read_parquet(rep_path)
            amt_col = next((c for c in rep.columns if "amount" in c.lower() and "split" not in c.lower()), None)
            if amt_col and "signature" in rep.columns:
                rep["_amt"] = pd.to_numeric(rep[amt_col], errors="coerce").fillna(0)
                grp = rep.groupby("signature")["_amt"].agg(["sum", "count"])
                sig_to_amt = grp["sum"].to_dict()
                sig_to_rows = grp["count"].to_dict()
        except Exception as e:
            print(f"  ⚠️ failed to load kpi_report for amount estimate: {e}")

    # Add amount/rows to problematic
    for p in problematic:
        sig = p["signature"]
        # Try exact match first; else search for sigs containing the pattern
        amt = sig_to_amt.get(sig, 0)
        nrows = sig_to_rows.get(sig, 0)
        p["amt"] = amt
        p["rows"] = nrows

    # Sort by abs amount (high impact first)
    problematic.sort(key=lambda x: -abs(x.get("amt", 0)))

    print(f"\nProblematic H_OTHER sigs matching new rule patterns: {len(problematic):,}")
    if sig_to_amt:
        total_amt = sum(abs(p.get("amt", 0)) for p in problematic)
        print(f"  Total |amount| impact: {total_amt:,.0f} MOP")
    print(f"  {'idx':>5}  {'pattern':<22}  {'expected':<14}  {'rows':>6}  {'amount':>12}  signature")
    print("  " + "-" * 130)
    for p in problematic[:50]:
        amt_str = f"{p.get('amt', 0):,.0f}" if p.get('amt') else "?"
        rows_str = f"{p.get('rows', 0):,}" if p.get('rows') else "?"
        sig_preview = p['signature'][:70]
        print(f"  {p['row_idx']:>5}  {p['pattern']:<22}  {p['expected_h']:<14}  {rows_str:>6}  {amt_str:>12}  {sig_preview}")

    if len(problematic) > 50:
        print(f"  ... +{len(problematic) - 50} more")

    if args.auto_fix:
        print(f"\n[auto-fix] Updating {len(problematic)} rows in feedback.xlsx...")
        for p in problematic:
            fb.at[p["row_idx"], "correct_horizontal"] = p["expected_h"]
            if "notes" in fb.columns:
                old_notes = str(fb.at[p["row_idx"], "notes"]) if pd.notna(fb.at[p["row_idx"], "notes"]) else ""
                fb.at[p["row_idx"], "notes"] = f"[auto-fix-r6] {old_notes}".strip()
        backup = fb_path.with_suffix(".backup.xlsx")
        if not backup.exists():
            pd.read_excel(fb_path).to_excel(backup, index=False)
            print(f"  backup saved: {backup.name}")
        fb.to_excel(fb_path, index=False)
        print(f"  ✅ updated {fb_path.name}")
    elif args.clear_only:
        print(f"\n[clear-only] Clearing {len(problematic)} H_OTHER overrides...")
        for p in problematic:
            fb.at[p["row_idx"], "correct_horizontal"] = ""
        backup = fb_path.with_suffix(".backup.xlsx")
        if not backup.exists():
            pd.read_excel(fb_path).to_excel(backup, index=False)
            print(f"  backup saved: {backup.name}")
        fb.to_excel(fb_path, index=False)
        print(f"  ✅ cleared in {fb_path.name}")
    else:
        print(f"\nDry-run only. To fix, add --auto-fix (set to expected_h) or --clear-only (revert to LLM/rules).")


if __name__ == "__main__":
    main()
