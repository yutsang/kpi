"""Step 4: see nodes.py docstring."""
from functools import partial, update_wrapper

from kedro.pipeline import Pipeline, node

from .nodes import step4_apply_tags


def _wrap(fn, *args, **kwargs):
    p = partial(fn, *args, **kwargs)
    update_wrapper(p, fn)
    return p

def create_pipeline(entity_key: str, **_) -> Pipeline:
    return Pipeline([
        node(
            func=_wrap(step4_apply_tags, entity_key),
            inputs=f"{entity_key}_step3_done",
            outputs=f"{entity_key}_step4_done",
            name=f"{entity_key}_step4_apply_tags",
        ),
    ])
