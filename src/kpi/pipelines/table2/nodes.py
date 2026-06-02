"""Kedro node entry-points for the table2 pipeline.

Both nodes always run from scratch (no skip-if-exists guard) — the user
explicitly asked for fresh builds since the data is small and local.
"""
from __future__ import annotations

from . import _logic


def table2_extract() -> str:
    _logic.extract_main()
    return "table2_extract_done"


def table2_analyze(_previous: str | None = None) -> str:
    _logic.analyze_main()
    return "table2_analyze_done"
