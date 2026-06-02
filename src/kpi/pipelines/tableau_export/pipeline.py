"""`tableau` delivery pipeline — standalone, Tableau files only.

  kedro run --pipeline=tableau

One node, loops all 6 entities → data/tableau. Reads step5 kpi_report.parquet — no re-classify.
"""
from kedro.pipeline import Pipeline, node

from .nodes import tableau_export


def create_pipeline(**_) -> Pipeline:
    return Pipeline([
        node(
            func=tableau_export,
            inputs=None,
            outputs="tableau_done",
            name="tableau",
        ),
    ])
