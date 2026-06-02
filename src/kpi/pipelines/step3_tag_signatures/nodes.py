"""Kedro node entry-point for step3_tag_signatures.
Logic lives in _logic.py; this file wraps it with progress tracking.
"""
from __future__ import annotations

from kpi.pipelines.common import wrap_with_progress

from . import _logic


def step3_tag_signatures(entity_key: str, _previous: str | None = None) -> str:
    return wrap_with_progress("step3_tag_signatures", entity_key, _logic.main)
