"""`tableau` pipeline — build ONLY the Tableau files for all 6 entities.

  kedro run --pipeline=tableau

  data/tableau/tableau_{yr}_{ent}.xlsx   (prep_tableau_25.run)

Reads kpi_report.parquet (step5 output) — no re-classify. Use this when you only need to refresh
Tableau; the full delivery bundle (投資方向 + tableau) is `generate`.
"""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[4]
_SCRIPTS = _ROOT / "scripts"
if str(_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS))


def main():
    import prep_tableau_25 as tab   # noqa: E402  (scripts/ on sys.path)

    print("=== tableau → data/tableau ===", flush=True)
    tab.run("per-entity-xlsx", "data/tableau")
    print("=== tableau done ===", flush=True)
