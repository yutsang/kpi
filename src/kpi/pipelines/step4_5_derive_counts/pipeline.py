"""Step 4.5: see nodes.py docstring."""
from functools import partial, update_wrapper

from kedro.pipeline import Pipeline, node

from .nodes import step4_5_derive_counts


def _wrap(fn, *args, **kwargs):
    p = partial(fn, *args, **kwargs)
    update_wrapper(p, fn)
    return p


def create_pipeline(entity_key: str, **_) -> Pipeline:
    return Pipeline([
        node(
            func=_wrap(step4_5_derive_counts, entity_key),
            inputs=f"{entity_key}_step4_done",
            outputs=f"{entity_key}_step4_5_done",
            name=f"{entity_key}_step4_5_derive_counts",
        ),
    ])
