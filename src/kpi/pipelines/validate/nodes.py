"""Kedro node entry-point for the validate pipeline."""
from __future__ import annotations

from kpi.pipelines.common import wrap_with_progress

from . import _logic


def validate(entity_key: str, _previous: str | None = None) -> str:
    return wrap_with_progress("validate", entity_key, _logic.main)
