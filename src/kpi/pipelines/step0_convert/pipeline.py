"""Step 0: convert source xlsx to parquet."""
from functools import partial, update_wrapper

from kedro.pipeline import Pipeline, node

from .nodes import step0_convert


def _wrap(fn, *args, **kwargs):
    p = partial(fn, *args, **kwargs)
    update_wrapper(p, fn)
    return p

def create_pipeline(entity_key: str, **_) -> Pipeline:
    return Pipeline([
        node(
            func=_wrap(step0_convert, entity_key),
            inputs=None,
            outputs=f"{entity_key}_step0_done",
            name=f"{entity_key}_step0_convert",
        ),
    ])
