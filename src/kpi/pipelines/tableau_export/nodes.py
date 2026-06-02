"""Kedro node entry-point for the `tableau` (tableau-only delivery) pipeline."""
from __future__ import annotations

from . import _logic


def tableau_export(_previous: str | None = None) -> str:
    _logic.main()
    return "tableau_done"
