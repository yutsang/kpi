"""Validate pipeline — runs ONLY when invoked as
`kedro run --pipeline=<entity>_validate`.

Structural cross-validation of LLM tagging quality. NOT in default 6-step.
"""
from functools import partial, update_wrapper

from kedro.pipeline import Pipeline, node

from .nodes import validate


def _wrap(fn, *args, **kwargs):
    p = partial(fn, *args, **kwargs)
    update_wrapper(p, fn)
    return p

def create_pipeline(entity_key: str, **_) -> Pipeline:
    return Pipeline([
        node(
            func=_wrap(validate, entity_key),
            inputs=None,
            outputs=f"{entity_key}_validate_done",
            name=f"{entity_key}_validate",
        ),
    ])
