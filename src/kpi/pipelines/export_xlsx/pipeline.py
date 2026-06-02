"""Export pipeline — runs ONLY when invoked as `kedro run --pipeline=<entity>_export`.

Reads <entity>_tagged_rows.parquet (step 4 output) and writes an xlsx (or csv)
with all rows sorted by abs(amount) desc, for manual review / sharing.

NOT included in the default 6-step pipeline.
"""
from functools import partial, update_wrapper

from kedro.pipeline import Pipeline, node

from .nodes import export_xlsx


def _wrap(fn, *args, **kwargs):
    p = partial(fn, *args, **kwargs)
    update_wrapper(p, fn)
    return p

def create_pipeline(entity_key: str, **_) -> Pipeline:
    return Pipeline([
        node(
            func=_wrap(export_xlsx, entity_key),
            inputs=None,
            outputs=f"{entity_key}_export_xlsx_done",
            name=f"{entity_key}_export_xlsx",
        ),
    ])
