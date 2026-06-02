"""Show-summary pipeline — runs ONLY when invoked as
`kedro run --pipeline=<entity>_summary`.

Reads <entity>_unique_signatures.xlsx (step 3 output) and prints summary stats
to console. NOT included in the default 6-step pipeline.
"""
from functools import partial, update_wrapper

from kedro.pipeline import Pipeline, node

from .nodes import show_summary


def _wrap(fn, *args, **kwargs):
    p = partial(fn, *args, **kwargs)
    update_wrapper(p, fn)
    return p

def create_pipeline(entity_key: str, **_) -> Pipeline:
    return Pipeline([
        node(
            func=_wrap(show_summary, entity_key),
            inputs=None,
            outputs=f"{entity_key}_show_summary_done",
            name=f"{entity_key}_show_summary",
        ),
    ])
