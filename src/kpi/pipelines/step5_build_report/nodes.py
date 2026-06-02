"""Kedro node entry-point for step5_build_report.
Logic lives in _logic.py; this file wraps it with progress tracking.
"""
from __future__ import annotations

from kpi.pipelines.common import wrap_with_progress

from . import _logic


def step5_build_report(entity_key: str, _previous: str | None = None) -> str:
    return wrap_with_progress("step5_build_report", entity_key, _logic.main)
