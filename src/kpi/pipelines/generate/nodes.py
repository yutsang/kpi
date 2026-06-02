"""Kedro node entry-point for the `generate` (delivery) pipeline."""
from __future__ import annotations

from . import _logic


def generate(_previous: str | None = None) -> str:
    _logic.main()
    return "generate_done"
