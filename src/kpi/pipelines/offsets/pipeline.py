"""Offsets pipeline — runs ONLY when invoked as
`kedro run --pipeline=<entity>_offsets`.

Reads <entity>_tagged_rows.parquet and outputs <entity>_offsets_review.xlsx
with rows tagged by which offset/reversal detection layer caught them. Designed
to support audit voucher sampling: when an auditor picks a sample row, they can
look it up in the xlsx and discover its paired/offsetting mate(s).

NOT included in the default 6-step pipeline.
"""
from functools import partial, update_wrapper

from kedro.pipeline import Pipeline, node

from .nodes import offsets


def _wrap(fn, *args, **kwargs):
    p = partial(fn, *args, **kwargs)
    update_wrapper(p, fn)
    return p

def create_pipeline(entity_key: str, **_) -> Pipeline:
    return Pipeline([
        node(
            func=_wrap(offsets, entity_key),
            inputs=None,
            outputs=f"{entity_key}_offsets_done",
            name=f"{entity_key}_offsets",
        ),
    ])
