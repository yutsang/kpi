"""Kedro node wrapper for step0_prebuild."""
from __future__ import annotations

from kpi.pipelines.common import wrap_with_progress

from . import _logic


def step0_prebuild(entity_key: str, _previous: str | None = None) -> str:
    return wrap_with_progress("step0_prebuild", entity_key, _logic.main)
