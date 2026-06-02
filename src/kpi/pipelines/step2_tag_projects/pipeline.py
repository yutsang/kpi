"""Step 2: see nodes.py docstring."""
from functools import partial, update_wrapper

from kedro.pipeline import Pipeline, node

from .nodes import step2_tag_projects


def _wrap(fn, *args, **kwargs):
    p = partial(fn, *args, **kwargs)
    update_wrapper(p, fn)
    return p

def create_pipeline(entity_key: str, **_) -> Pipeline:
    return Pipeline([
        node(
            func=_wrap(step2_tag_projects, entity_key),
            inputs=f"{entity_key}_step1_done",
            outputs=f"{entity_key}_step2_done",
            name=f"{entity_key}_step2_tag_projects",
        ),
    ])
