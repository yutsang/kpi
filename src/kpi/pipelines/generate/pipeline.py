"""`generate` delivery pipeline — standalone (not per-entity, not in the default run).

  kedro run --pipeline=generate

One node, loops all 6 entities: 投資方向 (pivot + 大表) → data/review  +  Tableau → data/tableau.
For a single artifact use the `tableau` (tableau-only) or `audit` pipeline. No re-classify.
"""
from kedro.pipeline import Pipeline, node

from .nodes import generate


def create_pipeline(**_) -> Pipeline:
    return Pipeline([
        node(
            func=generate,
            inputs=None,
            outputs="generate_done",
            name="generate",
        ),
    ])
