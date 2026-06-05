"""Step 0: convert source xlsx to parquet (skip if already done)."""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

from kpi.lib.conf import load_config, path  # noqa: E402
from kpi.lib.io_setup import force_unbuffered_io  # noqa: E402

force_unbuffered_io()

# Only these columns are mandatory. Others (system / submit_no / dicj_code etc.)
# are optional and only used by some entities; missing → ignored downstream.
REQUIRED_COL_KEYS = [
    "amount",
    "project",
    "ng11_category",
    "account_code",
    "account_desc",
    "description",
]


def _norm_match(x) -> str:
    """Normalise a value for equality comparison. Drops trailing '.0' from
    numeric-looking strings so YAML int 2025 matches Excel float 2025.0."""
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.endswith(".0") and s[:-2].lstrip("-").isdigit():
        return s[:-2]
    return s


def _apply_row_filter(df: pd.DataFrame, rules: dict, mode: str) -> pd.DataFrame:
    """Apply exclude_where or include_where rules.

    mode='exclude': drop rows where col IN values.
    mode='include': keep rows where col IN values (drop all others).
    """
    for col, vals in (rules or {}).items():
        if col not in df.columns:
            print(f"  WARNING: {mode}_where column '{col}' not in xlsx (ignored)")
            continue
        if not isinstance(vals, list):
            vals = [vals]
        str_vals = [_norm_match(v) for v in vals]
        before = len(df)
        mask = df[col].map(_norm_match).isin(str_vals)
        if mode == "exclude":
            df = df[~mask]
        else:
            df = df[mask]
        print(f"  {mode}_where {col} in {vals}: {before:,} → {len(df):,} rows")
    return df


def _hash_row(row, cols: dict) -> str:
    """Stable short hash from key fields. Used as fallback when no unique_id column."""
    import hashlib
    key_parts = [
        str(row.get(cols.get("account_code", ""), "")),
        str(row.get(cols.get("description", ""), "")),
        str(row.get(cols.get("amount", ""), "")),
        str(row.get(cols.get("posting_date", ""), "")),
        str(row.get(cols.get("project", ""), "")),
        str(row.get("_row_id", "")),     # final tiebreaker on identical rows
    ]
    return hashlib.sha256("|".join(key_parts).encode("utf-8")).hexdigest()[:16]


def _combine_cols(df: pd.DataFrame, source_cols: list[str], target_col: str, label: str) -> pd.DataFrame:
    """Concatenate source_cols (space-joined, blanks skipped) into target_col."""
    available = [c for c in source_cols if c in df.columns]
    if not available:
        print(f"  WARNING: {label}_cols {source_cols!r} — none found in xlsx, keeping '{target_col}' unchanged")
        return df
    df[target_col] = (
        df[available]
        .fillna("")
        .astype(str)
        .apply(lambda row: " ".join(p for p in row if p.strip()), axis=1)
        .str.strip()
    )
    print(f"  Combined {label} from {available} → '{target_col}'")
    return df


def _apply_col_combinations(df: pd.DataFrame, cols: dict) -> pd.DataFrame:
    """Apply project_name_cols / description_cols / account_code_cols / amount_cols
    to a single DataFrame using the given (possibly year-specific) cols dict."""
    project_name_cols = cols.get("project_name_cols") or []
    if project_name_cols:
        df = _combine_cols(df, project_name_cols, cols.get("project", "project"), "project")
    description_cols = cols.get("description_cols") or []
    if description_cols:
        df = _combine_cols(df, description_cols, cols.get("description", "description"), "description")
    account_code_cols = cols.get("account_code_cols") or []
    if account_code_cols:
        df = _combine_cols(df, account_code_cols, cols.get("account_code", "account_code"), "account_code")
    amount_cols = cols.get("amount_cols") or []
    if amount_cols:
        available = [c for c in amount_cols if c in df.columns]
        if not available:
            print(f"  WARNING: amount_cols {amount_cols!r} — none found, keeping '{cols.get('amount')}' unchanged")
        else:
            target = cols.get("amount", "amount")
            numeric = pd.DataFrame({
                c: pd.to_numeric(
                    df[c].astype(str).str.replace(",", "").str.strip(),
                    errors="coerce",
                ).fillna(0)
                for c in available
            })
            df[target] = numeric.sum(axis=1)
            print(f"  Combined amount = SUM{available} → '{target}'")
    return df


def _apply_ng_section_mapping(df: pd.DataFrame, cols: dict, ng_map: dict) -> pd.DataFrame:
    """Replace ng11_category column values per ng_section_mapping. Unmapped passes through.

    Useful for entities (e.g. MGM) whose section column has descriptive Chinese text
    rather than NG codes. Mapping example:
        ng_section_mapping:
          "Gaming Investment": "NG0"
          "娛樂表演 - 活動": "NG3"
    """
    if not ng_map:
        return df
    ng_col = cols.get("ng11_category", "")
    if not ng_col or ng_col not in df.columns:
        return df
    original = df[ng_col].astype(str).str.strip()
    mapped = original.map(ng_map)
    n_mapped = int(mapped.notna().sum())
    df[ng_col] = mapped.fillna(original)
    print(f"  Applied ng_section_mapping: {n_mapped:,}/{len(df):,} rows mapped")
    return df


def _apply_section_inference_by_capex_opex(df: pd.DataFrame, cols: dict, mapping: dict) -> pd.DataFrame:
    """Fill empty ng11_category rows based on capex_opex column value.

    For entities (e.g. MGM 2023/2024 historical) whose section column is empty for many
    rows but the wd / capex_opex column reliably implies a section.
    Mapping example:
        section_inference_by_capex_opex:
          Gaming_OPEX: NG0
          WD5_Patron: NG1
    Only fills rows where section is currently empty — never overwrites a real value.
    """
    if not mapping:
        return df
    ng_col = cols.get("ng11_category", "")
    co_col = cols.get("capex_opex", "")
    if not ng_col or ng_col not in df.columns or not co_col or co_col not in df.columns:
        return df
    sec = df[ng_col].astype(str).str.strip()
    empty_mask = sec.eq("") | sec.str.lower().eq("nan") | df[ng_col].isna()
    co = df[co_col].astype(str).str.strip()
    total_filled = 0
    for co_val, ng_val in mapping.items():
        m = empty_mask & co.eq(str(co_val))
        n = int(m.sum())
        if n:
            df.loc[m, ng_col] = ng_val
            total_filled += n
            print(f"  section_inference: {co_col}={co_val!r} (empty section) → {ng_val}  ({n:,} rows)")
    if total_filled:
        print(f"  section_inference total: {total_filled:,} rows filled")
    return df


def _drop_section_header_leak(df: pd.DataFrame, cols: dict) -> pd.DataFrame:
    """Drop rows where ng11_category column literally equals 'Section' (header text leak)."""
    ng_col = cols.get("ng11_category", "")
    if not ng_col or ng_col not in df.columns:
        return df
    mask = df[ng_col].astype(str).str.strip().eq("Section")
    n = int(mask.sum())
    if n:
        df = df[~mask].copy()
        print(f"  Dropped {n} rows with literal '{ng_col}'=='Section' (header leak)")
    return df


def _norm_colname(s: object) -> str:
    """Normalize a column name for fuzzy matching: collapse all whitespace
    (incl. \\n, \\r, \\t, multiple spaces) to single space, then lowercase.
    Used so YAML 'Reported Amount\\n(MOP)' matches Excel's actual newline
    encoding (\\n vs \\r\\n) and similar whitespace variants."""
    import re as _re
    return _re.sub(r"\s+", " ", str(s)).strip().lower()


def _rename_to_root_cols(df: pd.DataFrame, year_cols: dict, root_cols: dict) -> pd.DataFrame:
    """If a year's raw column name differs from the root standard name, rename it.
    e.g. 2024 raw has 'Entry Account' but root 'account_code' = 'Account' → rename.
    Only scalar (non-list) values are renamed.
    Uses fuzzy whitespace-normalized matching to handle \\n vs \\r\\n in headers."""
    rename_map: dict[str, str] = {}
    # Build normalized lookup of actual df columns
    df_norm_to_actual = {_norm_colname(c): c for c in df.columns}

    for k, root_name in root_cols.items():
        if not isinstance(root_name, str) or not root_name:
            continue
        year_name = year_cols.get(k)
        if not isinstance(year_name, str) or not year_name:
            continue
        if year_name == root_name:
            continue
        if root_name in df.columns:
            # Root col present, but a differently-named override col may ALSO exist with the real
            # values (e.g. a file carrying both 'NG11 Category' [blank] and 'NG11 category' [filled]).
            # Coalesce the override into the root col where the root is blank (never overwrites).
            yn = year_name if year_name in df.columns else df_norm_to_actual.get(_norm_colname(year_name))
            if yn and yn != root_name and yn in df.columns:
                root_s = df[root_name].astype("string")
                blank = root_s.isna() | root_s.str.strip().isin(["", "nan", "None", "NaN"])
                n_fill = int((blank & df[yn].notna()).sum())
                if n_fill:
                    df[root_name] = root_s.mask(blank, df[yn].astype("string"))
                    print(f"  coalesced {yn!r} → {root_name!r} ({n_fill:,} blank root rows filled)")
            continue  # root col already present

        # Try exact match on year_name
        if year_name in df.columns:
            rename_map[year_name] = root_name
            continue

        # Fuzzy: normalize whitespace and look up
        year_norm = _norm_colname(year_name)
        if year_norm in df_norm_to_actual:
            actual = df_norm_to_actual[year_norm]
            if actual not in rename_map:  # avoid double-rename
                rename_map[actual] = root_name
                print(f"  fuzzy-matched col {actual!r} → {root_name!r} (yaml: {year_name!r})")
    if rename_map:
        df = df.rename(columns=rename_map)
        print(f"  Renamed cols → root: {rename_map}")
    return df


def main():
    cfg = load_config()
    out_dir = path(cfg, "interim_dir")
    out = out_dir / f"{cfg['company']['code']}_raw.parquet"

    if out.exists():
        print(f"  [skip] parquet already exists: {out.relative_to(out.parent.parent.parent)}")
        return

    # Detect multi-year mode
    master = cfg.get("_master") or {}
    entity_key = cfg.get("_entity") or cfg["company"]["code"]
    ent = (master.get("companies") or {}).get(entity_key, {})
    yearly_sources = ent.get("yearly_sources") or []

    root_cols = cfg.get("columns") or {}
    header_row = cfg.get("header_row", 0)
    raw_dir = path(cfg, "raw_xlsx").parent  # data/<alias>/raw/

    if yearly_sources:
        # ── Multi-year mode: read each year's file, normalize, concat ─────────────
        parts = []
        for ycfg in yearly_sources:
            year = ycfg["year"]
            raw_filename = ycfg["raw_filename"]
            raw_sheet = ycfg.get("raw_sheet", 0)
            # Per-source header_row override (e.g. Galaxy 23SY 'Combine' sheet
            # has a label row above schema → header_row=1).
            yr_header_row = ycfg.get("header_row", header_row)
            fpath = raw_dir / raw_filename
            print(f"=== Year {year}: {fpath.name} ({raw_sheet!r}, header_row={yr_header_row}) ===")
            if not fpath.exists():
                print(f"  [skip] {fpath} not found")
                continue

            df_y = pd.read_excel(fpath, sheet_name=raw_sheet, header=yr_header_row,
                                 dtype=object, engine="openpyxl")
            print(f"  rows={len(df_y):,}  cols={len(df_y.columns)}")

            # Year-specific filters
            df_y = _apply_row_filter(df_y, ycfg.get("exclude_where"), mode="exclude")
            df_y = _apply_row_filter(df_y, ycfg.get("include_where"), mode="include")

            # Build effective cols = root + per-year override
            year_cols = {**root_cols, **(ycfg.get("columns_override") or {})}

            # Apply column combinations using year_cols
            df_y = _apply_col_combinations(df_y, year_cols)

            # Rename year-specific raw column names → root standard names
            df_y = _rename_to_root_cols(df_y, year_cols, root_cols)

            # static_period: when source file has no period column (e.g. Galaxy
            # 23SY supplement file), set a constant value to the period column
            # so step0_5 year_split can route all rows to the intended bucket.
            static_period = ycfg.get("static_period")
            if static_period:
                period_col = root_cols.get("period", "")
                if period_col:
                    df_y[period_col] = static_period
                    print(f"  static_period: set {period_col}={static_period!r} for all {len(df_y):,} rows")
                else:
                    print(f"  [warn] static_period={static_period!r} configured but no 'period' col mapping")

            # static_posting_year: force Posting Year col value. Used when step0_5
            # routes by Posting Year (not Period) but file has unreliable/mixed values.
            # E.g. Galaxy 23SY: 9.ADR rows store account codes in Posting Year col —
            # override to "2023" so the file routes cleanly to 24_23SY bucket.
            static_posting_year = ycfg.get("static_posting_year")
            if static_posting_year is not None:
                py_col = root_cols.get("posting_year", "")
                if py_col:
                    df_y[py_col] = str(static_posting_year)
                    print(f"  static_posting_year: set {py_col}={static_posting_year!r} for all {len(df_y):,} rows")
                else:
                    print(f"  [warn] static_posting_year={static_posting_year!r} configured but no 'posting_year' col mapping")

            # Tag the year
            df_y["report_year"] = year
            parts.append(df_y)

        if not parts:
            raise SystemExit("yearly_sources configured but no files loaded successfully")
        df = pd.concat(parts, ignore_index=True)
        print(f"\nCombined {len(parts)} year(s): {len(df):,} rows")
        cols = root_cols

    else:
        # ── Legacy single-year mode ──────────────────────────────────────────────
        src = path(cfg, "raw_xlsx")
        if not src.exists():
            raise FileNotFoundError(f"Source file not found: {src}")

        # Parquet source (e.g. MGM pre-built actuals) — read + apply all post-load transforms.
        if Path(src).suffix.lower() == ".parquet":
            out_dir.mkdir(parents=True, exist_ok=True)
            df = pd.read_parquet(src)
            cols = root_cols
            n_in = len(df)
            df = _drop_section_header_leak(df, cols)
            df = _apply_ng_section_mapping(df, cols, ent.get("ng_section_mapping") or {})
            df = _apply_section_inference_by_capex_opex(
                df, cols, ent.get("section_inference_by_capex_opex") or {}
            )
            df.to_parquet(out, index=False)
            print(f"  Read parquet → {out.name}  rows={n_in:,} → {len(df):,}")
            return

        print(f"Reading {src.name} ...")
        _rs = cfg["paths"]["raw_sheet"]
        if isinstance(_rs, (list, tuple)):
            # multi-sheet: read each + concat (column-aligned; missing cols → NaN). e.g. MGM combine+adjustment
            _parts = []
            for _sh in _rs:
                _dp = pd.read_excel(src, sheet_name=_sh, header=header_row, dtype=object, engine="openpyxl")
                print(f"  sheet {_sh!r}: rows={len(_dp):,}  cols={len(_dp.columns)}")
                _parts.append(_dp)
            df = pd.concat(_parts, ignore_index=True, sort=False)
        else:
            df = pd.read_excel(src, sheet_name=_rs, header=header_row, dtype=object, engine="openpyxl")
        print(f"  rows={len(df):,}  cols={len(df.columns)}")

        df = _apply_row_filter(df, cfg.get("exclude_where"), mode="exclude")
        df = _apply_row_filter(df, cfg.get("include_where"), mode="include")

        cols = root_cols
        df = _apply_col_combinations(df, cols)

    # ── Post-load transforms (apply to xlsx legacy + yearly_sources combined) ────
    df = _drop_section_header_leak(df, cols)
    df = _apply_ng_section_mapping(df, cols, ent.get("ng_section_mapping") or {})
    df = _apply_section_inference_by_capex_opex(
        df, cols, ent.get("section_inference_by_capex_opex") or {}
    )

    # ── Required-column check ─────────────────────────────────────────────────────
    missing_required = []
    for k in REQUIRED_COL_KEYS:
        col_name = cols.get(k, "")
        if not col_name:
            missing_required.append(f"{k} (not set in config)")
        elif col_name not in df.columns:
            missing_required.append(f"{k} → '{col_name}' (not in xlsx)")
    if missing_required:
        print("ERROR: required columns missing or unmapped:")
        for m in missing_required:
            print(f"  - {m}")
        print(f"\nActual xlsx columns ({len(df.columns)}):")
        for c in df.columns:
            print(f"  - {c!r}")
        raise SystemExit("Fix config/config.yaml columns mapping for this company.")

    # ── Type coercions ────────────────────────────────────────────────────────────
    df["_row_id"] = range(len(df))

    # Unique ID: use raw column(s) if configured, else hash key fields. _uid carries through
    # all downstream steps so future "target" updates can upsert by uid.
    import hashlib
    uid_col = cols.get("unique_id")
    uid_cols = cols.get("unique_id_cols") or []
    if uid_cols:
        available = [c for c in uid_cols if c in df.columns]
        if not available:
            print(f"  WARNING: unique_id_cols {uid_cols!r} — none found, falling back to hash")
            df["_uid"] = df.apply(lambda r: _hash_row(r, cols), axis=1)
        else:
            df["_uid"] = (
                df[available]
                .fillna("")
                .astype(str)
                .apply(lambda row: "|".join(p.strip() for p in row if p.strip()), axis=1)
            )
            blank_mask = df["_uid"].eq("")
            if blank_mask.any():
                df.loc[blank_mask, "_uid"] = df.loc[blank_mask].apply(
                    lambda r: _hash_row(r, cols), axis=1
                )
            print(f"  unique_id from columns {available}  ({(~blank_mask).sum():,} from cols, {blank_mask.sum():,} hashed)")
    elif uid_col and uid_col in df.columns:
        df["_uid"] = df[uid_col].astype("string").fillna("").str.strip()
        # Fall back to hash for blank cells
        blank_mask = df["_uid"].eq("")
        if blank_mask.any():
            df.loc[blank_mask, "_uid"] = df.loc[blank_mask].apply(
                lambda r: _hash_row(r, cols), axis=1
            )
        print(f"  unique_id from column '{uid_col}'  ({(~blank_mask).sum():,} from col, {blank_mask.sum():,} hashed)")
    else:
        df["_uid"] = df.apply(lambda r: _hash_row(r, cols), axis=1)
        print(f"  unique_id auto-generated (sha256 short hash of key fields)")

    if cols.get("amount") in df.columns:
        df[cols["amount"]] = pd.to_numeric(df[cols["amount"]], errors="coerce")
    if cols.get("posting_date") and cols["posting_date"] in df.columns:
        df[cols["posting_date"]] = pd.to_datetime(df[cols["posting_date"]], errors="coerce")
    if cols.get("posting_year") and cols["posting_year"] in df.columns:
        df[cols["posting_year"]] = pd.to_numeric(df[cols["posting_year"]], errors="coerce").astype("Int64")

    # Force every remaining object/mixed-type column to string. Without this, pyarrow
    # tries to infer dtypes and chokes on columns that mix int-looking and float-looking
    # strings (e.g. "104" and "53.1" both as object).
    typed_cols = {cols.get(k) for k in ("amount", "posting_date", "posting_year") if cols.get(k)}
    for c in df.columns:
        if c in typed_cols:
            continue
        if df[c].dtype == object:
            df[c] = df[c].astype("string")

    out_dir.mkdir(parents=True, exist_ok=True)
    df.to_parquet(out, index=False)
    print(f"  Wrote {out.name}  ({out.stat().st_size / 1e6:.1f} MB)")


if __name__ == "__main__":
    main()
