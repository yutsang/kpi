"""Kedro node entry-point for step4_5_derive_counts.
Logic lives in _logic.py; this file wraps it with progress tracking.
"""
from __future__ import annotations

from kpi.pipelines.common import wrap_with_progress

from . import _logic


def step4_5_derive_counts(entity_key: str, _previous: str | None = None) -> str:
    return wrap_with_progress("step4_5_derive_counts", entity_key, _logic.main)
