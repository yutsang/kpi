"""Kedro node entry-point for the show_summary pipeline."""
from __future__ import annotations

from kpi.pipelines.common import wrap_with_progress

from . import _logic


def show_summary(entity_key: str, _previous: str | None = None) -> str:
    return wrap_with_progress("show_summary", entity_key, _logic.main)
