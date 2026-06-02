"""`generate` pipeline — the full delivery bundle for all 6 entities, in one go.

  kedro run --pipeline=generate

Produces:
  1. 投資方向 (pivot + 大表)  → data/review/{ent}_投資方向_{year}.xlsx   (build_master_audit_25.build)
  2. Tableau files            → data/tableau/tableau_{yr}_{ent}.xlsx       (prep_tableau_25.run)

Reads kpi_report.parquet (step5 output) — does NOT re-run classification, so it's cheap to re-run
for delivery. The SJM admin-comp inject + guest→V resolution live inside build() (same source of
truth as step6), so they apply here automatically too.

Sibling pipelines if you only want ONE artifact:
  kedro run --pipeline=tableau   → data/tableau only (prep_tableau_25.run)
  kedro run --pipeline=audit     → cross-company audit sample only
  kedro run --pipeline=<entity>  → that entity's 投資方向 (via step6, inside the main run)
"""
from __future__ import annotations

import sys
from pathlib import Path

import yaml

_ROOT = Path(__file__).resolve().parents[4]          # …/generate → pipelines → kpi → src → repo root
_SCRIPTS = _ROOT / "scripts"
if str(_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS))

ENTITIES = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
            "vml": "company_4", "melco": "company_5", "mgm": "company_6"}


def main():
    import build_master_audit_25 as bma   # noqa: E402  (scripts/ on sys.path)
    import prep_tableau_25 as tab          # noqa: E402

    cats = yaml.safe_load((_ROOT / "conf" / "base" / "categories.yml").read_text(encoding="utf-8"))

    print("=== generate 1/2: 投資方向 (pivot + 大表) → data/review ===", flush=True)
    for ent, com in ENTITIES.items():
        try:
            bma.build(ent, com, cats)
        except Exception as e:
            print(f"  ⚠ {ent} 投資方向 failed: {e}", flush=True)

    print("\n=== generate 2/2: Tableau → data/tableau ===", flush=True)
    try:
        tab.run("per-entity-xlsx", "data/tableau")
    except Exception as e:
        print(f"  ⚠ tableau failed: {e}", flush=True)

    print("\n=== generate done ===", flush=True)
