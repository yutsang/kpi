"""Kedro node entry-point for the export_xlsx pipeline."""
from __future__ import annotations

from kpi.pipelines.common import wrap_with_progress

from . import _logic


def export_xlsx(entity_key: str, _previous: str | None = None) -> str:
    return wrap_with_progress("export_xlsx", entity_key, _logic.main)
