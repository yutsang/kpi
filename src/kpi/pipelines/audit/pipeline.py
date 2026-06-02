"""Cross-company audit pipeline.

  kedro run --pipeline=audit

Reads every entity's tagged_rows.parquet / unique_signatures.xlsx /
unique_projects.xlsx (if present) and writes a single sample-review xlsx at
data/audit/audit_review_<YYYYMMDD>.xlsx — for prompt tuning + spot review.

NOT per-entity. The legacy `<alias>_audit` pipelines are removed.
"""
from kedro.pipeline import Pipeline, node

from .nodes import audit


def create_pipeline(**_) -> Pipeline:
    return Pipeline([
        node(
            func=audit,
            inputs=None,
            outputs="audit_done",
            name="cross_company_audit",
        ),
    ])
