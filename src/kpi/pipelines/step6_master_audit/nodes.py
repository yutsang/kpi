"""Kedro node entry-point for step6_master_audit (logic in _logic.py)."""
from __future__ import annotations

from kpi.pipelines.common import wrap_with_progress

from . import _logic


def step6_master_audit(entity_key: str, _previous: str | None = None) -> str:
    return wrap_with_progress("step6_master_audit", entity_key, _logic.main)
