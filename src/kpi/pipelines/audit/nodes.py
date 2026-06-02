"""Kedro node entry-point for the cross-company audit pipeline.

Always runs from scratch — produces a single Excel sampling rows / signatures /
projects from all 6 entities into data/audit/audit_review_<date>.xlsx.
"""
from __future__ import annotations

from . import _logic


def audit() -> str:
    _logic.main()
    return "audit_done"
