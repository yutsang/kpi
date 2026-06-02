"""Step 6: build the per-entity audit master Excel (runs at the end of the main pipeline).

  kedro run --pipeline=galaxy   →  … → step5 → step6 → data/review/{ent}_投資方向_{year}.xlsx

So the audit master is ALWAYS regenerated right after classification — no separate step to
forget. Reuses scripts/build_master_audit_25.build() (one source of truth). For SJM, build()
also injects the admin-comp rows into 1_投資方向pivot (added once the combined→V mapping is set).
"""
from __future__ import annotations

import sys
from pathlib import Path

import yaml

from kpi.lib.conf import load_config
from kpi.lib.io_setup import force_unbuffered_io

force_unbuffered_io()

_ROOT = Path(__file__).resolve().parents[4]
_SCRIPTS = _ROOT / "scripts"
if str(_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS))


def main():
    cfg = load_config()
    folder = cfg["company"]["folder"]      # e.g. galaxy
    com = cfg["company"]["code"]           # e.g. company_1

    import build_master_audit_25 as bma    # noqa: E402  (scripts/ on sys.path)
    cats = yaml.safe_load((_ROOT / "conf" / "base" / "categories.yml").read_text(encoding="utf-8"))
    out = bma.build(folder, com, cats)
    if out:
        for p in (out if isinstance(out, list) else [out]):
            print(f"  audit master → {p}", flush=True)
    else:
        print("  [skip] no kpi_report.parquet yet (run step5 first)", flush=True)
