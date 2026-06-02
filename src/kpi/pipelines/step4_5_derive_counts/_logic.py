"""Step 4.5: Derive H_COUNT (數量/次數) per vertical via distinct count of count_cols.

Reads tagged_rows.parquet from step4, applies the per-company `count_cols` from
`columns:` block of parameters.yml, and writes a small parquet with one row per
vertical: {vertical_id, count}.

Verticals with `count_strategy: none` or `manual` in categories.yml are skipped
(count = None / left for human to fill).

Output: data/<alias>/interim/<entity_key>_vertical_counts.parquet
"""
from __future__ import annotations

import logging
from pathlib import Path

import pandas as pd
import yaml

from kpi.lib.conf import load_config, path
from kpi.lib.io_setup import force_unbuffered_io

force_unbuffered_io()
log = logging.getLogger(__name__)


def main():
    cfg = load_config()
    company = cfg["company"]["code"]
    interim = path(cfg, "interim_dir")

    tagged_path = interim / f"{company}_tagged_rows.parquet"
    if not tagged_path.exists():
        raise FileNotFoundError(f"Run step4 first. Missing: {tagged_path}")

    cols = cfg["columns"]
    raw_count_cols = cols.get("count_cols")
    # Accept either string (single col) or list (1+ cols) form in YAML.
    if not raw_count_cols:
        count_cols = []
    elif isinstance(raw_count_cols, str):
        count_cols = [raw_count_cols]
    elif isinstance(raw_count_cols, (list, tuple)):
        count_cols = list(raw_count_cols)
    else:
        count_cols = []
    if not count_cols:
        print(f"  [skip] no count_cols configured for {company}")
        return

    df = pd.read_parquet(tagged_path)
    if "vertical_id" not in df.columns:
        raise ValueError(f"tagged_rows missing 'vertical_id' column: {tagged_path}")

    # Resolve count_cols against actual columns (warn on missing)
    avail = [c for c in count_cols if c in df.columns]
    missing = [c for c in count_cols if c not in df.columns]
    if missing:
        print(f"  WARNING: count_cols missing in tagged_rows: {missing}")
    if not avail:
        raise ValueError(
            f"None of count_cols {count_cols!r} exist in tagged_rows columns: "
            f"{list(df.columns)[:30]}..."
        )

    # Always compute distinct count for every vertical declared in categories.yml.
    # count_strategy field is retained as documentation but no longer gates counting.
    cats_path = _find_categories_yml()
    with open(cats_path, encoding="utf-8") as f:
        cats = yaml.safe_load(f)
    verticals = cats.get("verticals", []) or []

    rows = []
    for v in verticals:
        v_id = v["id"]
        strategy = v.get("count_strategy", "")
        sub = df[df["vertical_id"] == v_id]
        if len(sub) == 0:
            rows.append({"vertical_id": v_id, "count_strategy": strategy, "count": 0})
            continue
        n = sub.groupby(avail, dropna=False).ngroups
        rows.append({"vertical_id": v_id, "count_strategy": strategy, "count": int(n)})
        print(f"  {v_id} ({strategy or 'n/a'}): {n} distinct from {len(sub):,} tagged rows")

    out = pd.DataFrame(rows, columns=["vertical_id", "count_strategy", "count"])
    out_path = interim / f"{company}_vertical_counts.parquet"
    out.to_parquet(out_path, index=False)
    print(f"\nWrote {len(out)} verticals -> {out_path.name}  (count_cols={avail})")


def _find_categories_yml() -> Path:
    """Locate conf/base/categories.yml by walking up from this file."""
    here = Path(__file__).resolve()
    for parent in here.parents:
        candidate = parent / "conf" / "base" / "categories.yml"
        if candidate.exists():
            return candidate
    raise FileNotFoundError("conf/base/categories.yml not found")


if __name__ == "__main__":
    main()
