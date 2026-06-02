"""Kedro node entry-point for step0_5_split_year."""
from __future__ import annotations

from kpi.pipelines.common import wrap_with_progress

from . import _logic


def step0_5_split_year(entity_key: str, _previous: str | None = None) -> str:
    return wrap_with_progress("step0_5_split_year", entity_key, _logic.main)
