"""Table 2 pipeline — parse 表二 forms then LLM-analyze each project.

Usage:
  kedro run --pipeline=table2            # extract + analyze
  kedro run --pipeline=table2_extract    # only parse xlsx → _extracted.xlsx
  kedro run --pipeline=table2_analyze    # only run LLM (requires existing _extracted.xlsx)
"""
from kedro.pipeline import Pipeline, node

from .nodes import table2_analyze, table2_extract


def create_pipeline(**_) -> Pipeline:
    return Pipeline([
        node(
            func=table2_extract,
            inputs=None,
            outputs="table2_extract_done",
            name="table2_extract",
        ),
        node(
            func=table2_analyze,
            inputs="table2_extract_done",
            outputs="table2_analyze_done",
            name="table2_analyze",
        ),
    ])


def create_extract_only_pipeline(**_) -> Pipeline:
    return Pipeline([
        node(
            func=table2_extract,
            inputs=None,
            outputs="table2_extract_done",
            name="table2_extract",
        ),
    ])


def _analyze_no_input() -> str:
    """Analyze-only entry point (no upstream input)."""
    return table2_analyze()


def create_analyze_only_pipeline(**_) -> Pipeline:
    # Skip extract — caller must have already produced _extracted.xlsx.
    return Pipeline([
        node(
            func=_analyze_no_input,
            inputs=None,
            outputs="table2_analyze_done",
            name="table2_analyze",
        ),
    ])
