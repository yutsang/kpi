"""Non-destructive fix for 'new-year projects have no V': APPEND projects that exist in
{code}_raw.parquet but are missing from {code}_unique_projects.xlsx — nothing is deleted, no
re-extract, no step2/step3 re-run, sig file + LLM cache untouched. Existing rows (and their
manual_vertical/llm_vertical tags) are kept exactly as-is.

Use when an entity's interim was built before a year was added (e.g. Melco 24-only projects).
After top-up the new rows have blank V; supply it the normal way:
  python scripts/topup_projects.py --entity melco
  python scripts/dump_project_context.py --entity melco     # now covers the new projects
  # → Claude classifies → classify_vertical_from_ctx --write → inject_manual_vertical --entity melco
  kedro run --pipeline=melco                                 # step2 skips (already tagged), step4
                                                             # uses manual_vertical for the new ones
"""
from __future__ import annotations
import argparse, re, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception: pass
import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parent.parent
ENTITIES = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
            "vml": "company_4", "melco": "company_5", "mgm": "company_6"}


def norm(s) -> str:
    return re.sub(r"\s+", " ", str(s)).strip().lower()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", required=True, choices=sorted(ENTITIES))
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    code = ENTITIES[a.entity]
    interim = ROOT / "data" / a.entity / "interim"
    pq = interim / f"{code}_raw.parquet"
    px = interim / f"{code}_unique_projects.xlsx"
    if not pq.exists() or not px.exists():
        print(f"X missing {pq.name} or {px.name} under data/{a.entity}/interim"); return

    cols = yaml.safe_load((ROOT / "conf" / code / "parameters.yml").read_text(encoding="utf-8"))["columns"]
    pcol = cols["project"]

    proj_df = pd.read_excel(px)
    if pcol not in proj_df.columns:
        print(f"X project col {pcol!r} not in {px.name}; cols={list(proj_df.columns)}"); return
    have = set(proj_df[pcol].map(norm))

    raw_projs = pd.read_parquet(pq, columns=[pcol])[pcol].astype("string").fillna("")
    seen, missing = set(), []
    for p in raw_projs:
        k = norm(p)
        if not k or k == "nan" or k in have or k in seen:
            continue
        seen.add(k)
        missing.append(str(p))

    print(f"[{a.entity}] unique_projects has {len(proj_df):,} projects; "
          f"raw.parquet has {len(have) + len(missing):,} distinct; MISSING = {len(missing):,}")
    if not missing:
        print("  nothing to add — unique_projects already covers the parquet."); return
    for m in missing[:15]:
        print(f"      + {m[:90]}")
    if len(missing) > 15:
        print(f"      … +{len(missing) - 15} more")

    if a.dry_run:
        print("  (dry-run — not written)"); return
    new = pd.DataFrame({pcol: missing}).reindex(columns=proj_df.columns)
    out = pd.concat([proj_df, new], ignore_index=True)
    out.to_excel(px, index=False)
    print(f"  appended {len(missing):,} projects (blank V) → {px.name}. "
          f"Next: dump_project_context --entity {a.entity}")


if __name__ == "__main__":
    main()
