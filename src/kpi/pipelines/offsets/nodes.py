"""Kedro node entry-point for the offsets pipeline."""
from __future__ import annotations

from kpi.pipelines.common import wrap_with_progress

from . import _logic


def offsets(entity_key: str, _previous: str | None = None) -> str:
    return wrap_with_progress("offsets", entity_key, _logic.main)
