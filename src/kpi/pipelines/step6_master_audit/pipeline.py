"""Step 6: see nodes.py docstring. Chains after step5 in the main pipeline."""
from functools import partial, update_wrapper

from kedro.pipeline import Pipeline, node

from .nodes import step6_master_audit


def _wrap(fn, *args, **kwargs):
    p = partial(fn, *args, **kwargs)
    update_wrapper(p, fn)
    return p


def create_pipeline(entity_key: str, **_) -> Pipeline:
    return Pipeline([
        node(
            func=_wrap(step6_master_audit, entity_key),
            inputs=f"{entity_key}_step5_done",
            outputs=f"{entity_key}_step6_done",
            name=f"{entity_key}_step6_master_audit",
        ),
    ])
