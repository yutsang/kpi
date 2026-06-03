"""Force an entity's step0/step1 to rebuild from scratch — run this AFTER adding a new year
source (e.g. VML 2023), then `kedro run --pipeline=<ent>`.

Why it's needed: step0 skips if `<code>_raw.parquet` exists, and step1 skips if
`<code>_unique_signatures.xlsx` exists. So just adding a yearly_source does NOT re-ingest it
and does NOT add the new-year projects to unique_projects (they end up with no V). This deletes:
  data/<ent>/interim/<code>_raw.parquet         (→ step0 re-converts ALL sources incl new year)
  data/<ent>/interim/<code>_unique_*.xlsx       (→ step1 re-extracts projects/sigs/accts/vendors)
It does NOT touch the LLM cache (interim/cache/…) or the output kpi_report.

  python scripts/force_reextract.py --entity vml
  python scripts/force_reextract.py --entity vml --dry-run
Then: kedro run --pipeline=vml
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception: pass

ROOT = Path(__file__).resolve().parent.parent
ENTITIES = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
            "vml": "company_4", "melco": "company_5", "mgm": "company_6"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", required=True, choices=sorted(ENTITIES))
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    code = ENTITIES[a.entity]
    interim = ROOT / "data" / a.entity / "interim"
    if not interim.exists():
        print(f"X {interim} does not exist"); return

    targets = []
    pq = interim / f"{code}_raw.parquet"
    if pq.exists():
        targets.append(pq)
    targets += sorted(interim.glob(f"{code}_unique_*.xlsx"))
    if not targets:
        print(f"[{a.entity}] nothing to delete in {interim.relative_to(ROOT)} "
              f"(no {code}_raw.parquet / {code}_unique_*.xlsx)")
        return
    for t in targets:
        print(("would delete  " if a.dry_run else "deleted  ") + str(t.relative_to(ROOT)))
        if not a.dry_run:
            t.unlink()
    print(f"\n[{a.entity}] {'(dry-run) ' if a.dry_run else ''}{len(targets)} file(s). "
          f"Next: kedro run --pipeline={a.entity}")


if __name__ == "__main__":
    main()
