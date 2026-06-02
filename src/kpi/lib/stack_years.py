"""Stack prior-year raw files into the current year's raw parquet.

Usage:
  python -m kpi.lib.stack_years --entity company_3
  python -m kpi.lib.stack_years --entity mgm

Reads:
  conf/<entity>/parameters.yml → prior_raw_files list + raw_filename (current year)
  All files are expected in data/<alias>/raw/

Writes:
  data/<alias>/raw/<raw_filename>   (overwrites with stacked data)

The pipeline then runs unchanged — step0 sees a single parquet with all years,
step5 splits it by posting_year automatically.

For MGM (company_6): generate prior-year parquets from mgm_build.py first
(update F1–F4 filenames at the top of mgm_build.py to point to the old year's
Excel files, run it, then rename the output before running this script).
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

from kpi.lib.conf import load_config, load_master, resolve_entity_key, folder_for, path, _read_yaml, CONF_DIR
from kpi.lib.io_setup import force_unbuffered_io

force_unbuffered_io()


def main():
    cfg = load_config()
    company = cfg["company"]["code"]
    raw_dir = path(cfg, "interim_dir").parent / "raw"  # data/<alias>/raw/
    raw_dir.mkdir(parents=True, exist_ok=True)

    master = cfg["_master"]
    ent = (master.get("companies") or {}).get(company, {})
    prior_files = ent.get("prior_raw_files") or []

    current_filename = cfg["paths"]["raw_filename"]
    current_path = raw_dir / current_filename

    if not prior_files:
        print(f"No prior_raw_files configured for {company} — nothing to stack.")
        print(f"Add them to conf/{company}/parameters.yml under prior_raw_files:")
        print(f"  prior_raw_files:")
        print(f"    - filename: \"company_X_actuals_2023.xlsx\"")
        print(f"      raw_sheet: 0")
        return

    parts: list[pd.DataFrame] = []

    # Load prior year files
    for entry in prior_files:
        fname = entry.get("filename", "")
        if not fname:
            print(f"  [skip] empty filename in prior_raw_files entry", flush=True)
            continue
        fpath = raw_dir / fname
        if not fpath.exists():
            raise FileNotFoundError(
                f"Prior-year file not found: {fpath}\n"
                f"Place it in {raw_dir} before running this script."
            )
        print(f"Loading prior: {fpath.name} ...", flush=True)
        if fpath.suffix.lower() == ".parquet":
            df = pd.read_parquet(fpath)
        else:
            sheet = entry.get("raw_sheet", 0)
            header = entry.get("header_row", cfg.get("header_row", 0))
            df = pd.read_excel(fpath, sheet_name=sheet, header=header, dtype=object, engine="openpyxl")
            # Coerce amount column to numeric
            amt_col = cfg["columns"].get("amount", "")
            if amt_col and amt_col in df.columns:
                df[amt_col] = pd.to_numeric(df[amt_col], errors="coerce")
            # Coerce posting_date
            date_col = cfg["columns"].get("posting_date", "")
            if date_col and date_col in df.columns:
                df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
        print(f"  {len(df):,} rows  cols={len(df.columns)}", flush=True)
        parts.append(df)

    # Load current year file
    if current_path.exists():
        print(f"Loading current: {current_path.name} ...", flush=True)
        if current_path.suffix.lower() == ".parquet":
            df_cur = pd.read_parquet(current_path)
        else:
            sheet = cfg["paths"].get("raw_sheet", 0)
            df_cur = pd.read_excel(current_path, sheet_name=sheet, header=cfg.get("header_row", 0),
                                   dtype=object, engine="openpyxl")
        print(f"  {len(df_cur):,} rows  cols={len(df_cur.columns)}", flush=True)
        parts.append(df_cur)
    else:
        print(f"  [warn] current-year file not found: {current_path.name}", flush=True)

    if not parts:
        raise SystemExit("Nothing to stack. Check your file paths.")

    # Align columns — use union, fill missing with NA
    all_cols = list(dict.fromkeys(c for df in parts for c in df.columns))
    stacked = pd.concat(
        [df.reindex(columns=all_cols) for df in parts],
        ignore_index=True,
    )

    print(f"\nStacked: {len(stacked):,} rows total", flush=True)

    # Save — always as parquet (step0 handles parquet copy natively)
    out_name = current_filename if current_filename.endswith(".parquet") else Path(current_filename).stem + ".parquet"
    out_path = raw_dir / out_name
    stacked.to_parquet(out_path, index=False)
    print(f"Wrote: {out_path}  ({out_path.stat().st_size / 1e6:.1f} MB)", flush=True)

    # If raw_filename was xlsx, update the effective path to use the parquet
    if out_name != current_filename:
        print(
            f"\n  NOTE: raw_filename in parameters.yml is '{current_filename}' (xlsx).\n"
            f"  The stacked output was written as '{out_name}' (parquet).\n"
            f"  Update raw_filename in conf/{company}/parameters.yml to '{out_name}' before running the pipeline.",
            flush=True,
        )

    print("\nDone. Run the pipeline:")
    print(f"  kedro run --pipeline={cfg['company'].get('alias') or company}")


if __name__ == "__main__":
    main()
