"""Step 4: Apply project + signature tags to all 624k rows.

Reads:
  - data/interim/<company>_raw.parquet
  - data/interim/<company>_unique_projects.xlsx     (vertical tags)
  - data/interim/<company>_unique_signatures.xlsx   (horizontal tags)

For each row:
  1. Vertical (縱向) — looked up by project name. Priority: manual_vertical > llm_vertical.
  2. Horizontal (橫向) — looked up by signature. Priority: manual_horizontal > llm_horizontal.
  3. Capex/Opex — taken directly from source column (already clean).
  4. Row_type — derived from amount sign (negative → adjustment, else → normal).
  5. NG scope — 'gaming' if NG0, else 'non_gaming'.

Output: data/interim/<company>_tagged_rows.parquet
  Original columns + vertical_id / horizontal_id / labels / tag_sources / row_type / ng_scope.
"""
from __future__ import annotations

import gc
import sys
from pathlib import Path

import pandas as pd
from tqdm import tqdm

from kpi.lib.conf import load_config, load_categories, path  # noqa: E402
from kpi.lib.io_setup import force_unbuffered_io  # noqa: E402
from kpi.lib.text import compose_signature, normalize_description, resolve_signature_fields  # noqa: E402

force_unbuffered_io()


def _final_value(manual_col: pd.Series, llm_col: pd.Series) -> tuple[pd.Series, pd.Series]:
    """Return (final_value, source) where source is 'manual' if manual is set, else taken
    from the existing tag_source column passed alongside (handled by caller)."""
    manual = manual_col.astype("string").fillna("").str.strip()
    has_manual = manual.ne("") & manual.ne("nan")
    final = manual.where(has_manual, llm_col.astype("string").fillna(""))
    return final, has_manual


def _apply_row_overrides(df, rules, id_col, src_col, account_code_col, account_desc_col,
                         source_tag="rule:row"):
    """Apply conf-driven ROW-LEVEL overrides to ``id_col`` (first-match-wins per row).

    Needed for decisions that depend on a per-row helper column which a project- or
    signature-level lookup cannot capture (e.g. an internal-payroll flag, a comp
    nature column, or a project-nature column). Each rule is::

        {when: {<conditions, ALL must hold>}, set: <ID>}

    Supported conditions:
      account_code            exact match on the account_code column
      account_code_prefix     startswith on the account_code column
      account_desc            exact match on the account_desc column
      account_desc_contains   case-insensitive substring on account_desc
      column_equals   {col, value}      exact match on any raw column
      column_prefix   {col, value}      startswith on any raw column
      column_contains {col, any:[...]}  case-insensitive substring (any-of)
      column_nonempty <col>             the column has a non-blank value
      column_nonzero  <col>             the column's numeric value is non-blank AND != 0

    Rows set by an earlier rule are not touched again. Returns rows overridden.
    """
    if not rules:
        return 0
    acc = df[account_code_col].astype("string").fillna("").str.strip()
    ad = df[account_desc_col].astype("string").fillna("").str.strip()

    def _col(name):
        if name in df.columns:
            return df[name].astype("string").fillna("")
        return pd.Series("", index=df.index, dtype="string")

    done = pd.Series(False, index=df.index)
    n_total = 0
    for rule in rules:
        # column_map: set id_col from a {column_value: ID} dict (self-contained rule,
        # no separate `set`). Used to adopt a project-team category column wholesale.
        cm = rule.get("column_map")
        if cm:
            col = _col(cm["col"]).str.strip()
            for val, target in (cm.get("map") or {}).items():
                sub = (~done) & col.eq(str(val).strip())
                df.loc[sub, id_col] = target
                df.loc[sub, src_col] = source_tag
                done |= sub
                n_total += int(sub.sum())
            continue
        when = rule.get("when") or {}
        setval = rule.get("set")
        if not setval:
            continue
        mask = ~done
        if "account_code" in when:
            mask &= acc.eq(str(when["account_code"]).strip())
        if "account_code_prefix" in when:
            mask &= acc.str.startswith(str(when["account_code_prefix"]).strip())
        if "account_desc" in when:
            mask &= ad.eq(str(when["account_desc"]).strip())
        if "account_desc_contains" in when:
            mask &= ad.str.contains(str(when["account_desc_contains"]), case=False, na=False, regex=False)
        if "column_equals" in when:
            spec = when["column_equals"]
            mask &= _col(spec["col"]).str.strip().eq(str(spec["value"]).strip())
        if "column_prefix" in when:
            spec = when["column_prefix"]
            mask &= _col(spec["col"]).str.strip().str.startswith(str(spec["value"]).strip())
        if "column_contains" in when:
            spec = when["column_contains"]
            col = _col(spec["col"])
            sub = pd.Series(False, index=df.index)
            for kw in spec.get("any", []):
                sub |= col.str.contains(str(kw), case=False, na=False, regex=False)
            mask &= sub
        if "column_nonempty" in when:
            col = _col(when["column_nonempty"]).str.strip()
            mask &= ~col.isin(["", "nan", "NaN", "<NA>", "None", "NaT"])
        if "column_nonzero" in when:
            num = pd.to_numeric(_col(when["column_nonzero"]), errors="coerce")
            mask &= num.notna() & (num != 0)
        df.loc[mask, id_col] = setval
        df.loc[mask, src_col] = source_tag
        done |= mask
        n_total += int(mask.sum())
    return n_total


def main():
    cfg = load_config()
    cats = load_categories()
    company = cfg["company"]["code"]
    interim = path(cfg, "interim_dir")
    cols = cfg["columns"]

    raw_path = interim / f"{company}_raw.parquet"
    proj_xlsx = interim / f"{company}_unique_projects.xlsx"
    sig_xlsx = interim / f"{company}_unique_signatures.xlsx"
    for p in (raw_path, proj_xlsx, sig_xlsx):
        if not p.exists():
            raise FileNotFoundError(f"Missing prerequisite: {p}")

    print(f"Loading inputs ...", flush=True)
    df = pd.read_parquet(raw_path)
    proj_df = pd.read_excel(proj_xlsx)
    sig_df = pd.read_excel(sig_xlsx)
    print(f"  rows={len(df):,}  projects={len(proj_df)}  signatures={len(sig_df)}", flush=True)

    project_col = cols["project"]
    vlookup = {v["id"]: v for v in cats["verticals"]}
    hlookup = {h["id"]: h for h in cats["horizontals"]}

    # ----- Project → vertical lookup -----
    if "manual_vertical" not in proj_df.columns:
        proj_df["manual_vertical"] = ""
    if "llm_vertical" not in proj_df.columns:
        proj_df["llm_vertical"] = ""
    if "tag_source" not in proj_df.columns:
        proj_df["tag_source"] = ""
    final_v, has_manual_v = _final_value(proj_df["manual_vertical"], proj_df["llm_vertical"])
    proj_df["_final_vertical"] = final_v
    proj_df["_vertical_source"] = has_manual_v.map({True: "manual", False: ""}).where(
        has_manual_v, proj_df["tag_source"].astype("string").fillna("")
    )
    proj_to_v = dict(zip(proj_df[project_col].astype("string").fillna(""), proj_df["_final_vertical"]))
    proj_to_v_src = dict(zip(proj_df[project_col].astype("string").fillna(""), proj_df["_vertical_source"]))

    # ----- Signature → horizontal lookup -----
    if "manual_horizontal" not in sig_df.columns:
        sig_df["manual_horizontal"] = ""
    if "llm_horizontal" not in sig_df.columns:
        sig_df["llm_horizontal"] = ""
    if "tag_source" not in sig_df.columns:
        sig_df["tag_source"] = ""
    final_h, has_manual_h = _final_value(sig_df["manual_horizontal"], sig_df["llm_horizontal"])
    sig_df["_final_horizontal"] = final_h
    sig_df["_horizontal_source"] = has_manual_h.map({True: "manual", False: ""}).where(
        has_manual_h, sig_df["tag_source"].astype("string").fillna("")
    )
    sig_to_h = dict(zip(sig_df["signature"].astype("string"), sig_df["_final_horizontal"]))
    sig_to_h_src = dict(zip(sig_df["signature"].astype("string"), sig_df["_horizontal_source"]))

    # ----- Per-row signature recomputation (must match step 1 exactly) -----
    print("Building per-row signatures (this is the slow part) ...", flush=True)
    _ent_cfg = ((cfg.get("_master") or {}).get("companies") or {}).get(company, {}) or {}
    _fields = resolve_signature_fields(_ent_cfg, cols, df.columns)
    acc = df[cols["account_code"]].astype("string").fillna("").str.strip()
    ad = df[cols["account_desc"]].astype("string").fillna("").str.strip()
    desc_raw = df[cols["description"]]
    tqdm.pandas(desc="normalize description", file=sys.stderr)
    desc_norm = desc_raw.progress_apply(normalize_description)
    # job_code component (empty Series when not configured); helper picks it up only if in _fields.
    jc_col = cols.get("job_code", "")
    jc = (df[jc_col].astype("string").fillna("").str.strip()
          if jc_col and jc_col in df.columns else pd.Series([""] * len(df), index=df.index))
    # Same fields + shared helper as step1 build_signatures() → byte-identical signature.
    df["signature"] = compose_signature(
        {"account_code": acc, "account_desc": ad, "desc_norm": desc_norm, "job_code": jc}, _fields)
    # Free the large per-row string intermediates immediately — on memory-tight boxes
    # (Galaxy 1.15M rows) keeping acc/ad/desc_norm alive alongside the new signature column
    # plus the full frame exhausts RAM at write time.
    del acc, ad, desc_raw, desc_norm, jc
    gc.collect()

    # ----- Apply lookups -----
    print("Joining vertical + horizontal tags ...", flush=True)
    proj_str = df[project_col].astype("string").fillna("")
    df["vertical_id"] = proj_str.map(proj_to_v)
    df["vertical_source"] = proj_str.map(proj_to_v_src)

    df["horizontal_id"] = df["signature"].map(sig_to_h)
    df["horizontal_source"] = df["signature"].map(sig_to_h_src)

    # ----- Row-level overrides (conf-driven) -----
    # Decisions that depend on a per-row helper column the project/signature lookup
    # cannot capture (internal-payroll flag, comp nature, project nature, ...).
    # load_config() only surfaces a fixed set of keys on cfg; the per-entity blocks
    # live under cfg["_master"]["companies"][<code>] (same access step3 uses).
    _ent_cfg = ((cfg.get("_master") or {}).get("companies") or {}).get(company, {}) or {}
    n_rv = _apply_row_overrides(df, _ent_cfg.get("row_vertical_overrides") or [],
                                "vertical_id", "vertical_source",
                                cols["account_code"], cols["account_desc"])
    n_rh = _apply_row_overrides(df, _ent_cfg.get("row_horizontal_overrides") or [],
                                "horizontal_id", "horizontal_source",
                                cols["account_code"], cols["account_desc"])
    if n_rv or n_rh:
        print(f"Row-level overrides applied: vertical={n_rv:,} rows  horizontal={n_rh:,} rows", flush=True)

    # ----- Labels (after overrides) -----
    df["vertical_label"] = df["vertical_id"].map(lambda v: vlookup.get(v, {}).get("label", "") if v else "")
    df["horizontal_label"] = df["horizontal_id"].map(lambda h: hlookup.get(h, {}).get("label", "") if h else "")

    # ----- Capex/Opex from source (normalize casing: "opex"→"Opex", "CAPEX"→"Capex") -----
    # capex_opex_unmatched (per-entity): when the source col carries codes that aren't capex/opex
    # (e.g. MGM Source = CAPEX / WD1 / WD5-Patron / ADJUSTMENT), map every non-capex value to it.
    _co_unmatched = _ent_cfg.get("capex_opex_unmatched") if isinstance(_ent_cfg, dict) else None
    def _norm_co(s: str) -> str:
        ls = s.strip().lower()
        if ls == "capex" or ls.startswith("capital"):
            return "Capex"
        if ls == "opex" or ls.startswith("operat"):
            return "Opex"
        return _co_unmatched if _co_unmatched else s
    df["final_capex_opex"] = df[cols["capex_opex"]].astype("string").fillna("").map(_norm_co)

    # ----- 維護費 only counts as opex: a CAPEX 維護 row belongs to 建設與設施支出 -----
    # H_MAINTENANCE is opex-scoped; when the same row is capex it is a 翻新/結構性 spend, so
    # reclassify it to H_CONSTRUCTION (label updated too).
    _maint_capex = df["horizontal_id"].eq("H_MAINTENANCE") & df["final_capex_opex"].eq("Capex")
    if (_ent_cfg.get("maint_capex_to_construction", True) if isinstance(_ent_cfg, dict) else True) and _maint_capex.any():
        df.loc[_maint_capex, "horizontal_id"] = "H_CONSTRUCTION"
        df.loc[_maint_capex, "horizontal_label"] = hlookup.get("H_CONSTRUCTION", {}).get("label", "建設與設施支出")
        print(f"維護費(capex) → 建設與設施支出: {int(_maint_capex.sum()):,} rows", flush=True)

    # ----- capex 只能係已資本化嘅 construction/equipment（+ 人工 capex 可單獨拆出）-----
    # 會計上 capex = 入帳資產，唔應該係 專業服務費/廣告/合約 等 opex-類 H。任何 capex 行但
    # horizontal_id 唔喺 allow set → 撥去 `to`（預設 H_CONSTRUCTION）。conf-gated：只有設咗
    # `capex_force` 嘅 entity（SJM）先會行，其他（Melco capex-維護費 等）唔受影響。
    _cf = (_ent_cfg.get("capex_force") or {}) if isinstance(_ent_cfg, dict) else {}
    if _cf:
        _allow = set(_cf.get("allow") or ["H_CONSTRUCTION", "H_EQUIP", "H_LABOR"])
        _to = _cf.get("to", "H_CONSTRUCTION")
        _force = (df["final_capex_opex"].eq("Capex") & df["horizontal_id"].notna()
                  & df["horizontal_id"].ne("") & ~df["horizontal_id"].isin(_allow))
        if _force.any():
            print(f"capex_force → {_to}: {int(_force.sum()):,} rows "
                  f"(from {df.loc[_force,'horizontal_id'].value_counts().to_dict()})", flush=True)
            df.loc[_force, "horizontal_id"] = _to
            df.loc[_force, "horizontal_label"] = hlookup.get(_to, {}).get("label", "建設與設施支出")

    # ----- NG scope (gaming vs non-gaming) -----
    # Accept both code ("NG0") and Chinese label ("博彩") for gaming rows,
    # since different companies store the ng11_category in different formats.
    ng_str = df[cols["ng11_category"]].astype("string").fillna("")
    is_gaming = ng_str.eq("NG0") | ng_str.str.startswith("博彩")
    df["ng_scope"] = is_gaming.map({True: "gaming", False: "non_gaming"})

    # ----- Row type from amount sign -----
    amt = pd.to_numeric(df[cols["amount"]], errors="coerce")
    df["row_type"] = amt.apply(lambda x: "adjustment" if pd.notna(x) and x < 0 else "normal")

    # ----- Save -----
    # Reclaim RAM before the write. The signature column (1.15M concatenated strings) is the
    # single biggest hog and is NOT read by any downstream step (step4_5/5/6/generate/audit) —
    # drop it. Also free every lookup dict + input frame still pinned in scope. Without this the
    # frame + the parquet conversion don't fit on memory-tight boxes (Galaxy OOM'd even on the
    # chunk fallback's 781 KiB alloc because all of these were still resident).
    df.drop(columns=["signature"], inplace=True, errors="ignore")
    del proj_to_v, proj_to_v_src, sig_to_h, sig_to_h_src
    del proj_df, sig_df, proj_str, ng_str, amt
    del final_v, has_manual_v, final_h, has_manual_h
    # pandas 'string' dtype → object before the parquet write: StringArray.__arrow_array__ is
    # memory-heavy AND spawns a thread pool, so on a RAM-exhausted box it hits 'realloc failed' /
    # 'can't start new thread'. object converts via the lean numpy path, single-threaded.
    for _c in df.columns:
        if str(df[_c].dtype) == "string":
            df[_c] = df[_c].astype(object)
    gc.collect()

    out_path = interim / f"{company}_tagged_rows.parquet"
    try:
        df.to_parquet(out_path, index=False)
    except Exception as _e:
        # Low-memory fallback (big entities like Galaxy 1.15M rows): convert + write in small row
        # chunks so pyarrow never builds the whole Table at once, freeing each chunk immediately.
        # nthreads=1 avoids the thread-pool spawn that fails ('can't start new thread') when RAM
        # is exhausted.
        print(f"  [save] to_parquet failed ({_e}); retrying with chunked low-memory write…", flush=True)
        import pyarrow as pa
        import pyarrow.parquet as pq
        _CH = 50_000
        _schema = pa.Table.from_pandas(df.iloc[:2000], preserve_index=False, nthreads=1).schema
        _writer = pq.ParquetWriter(out_path, _schema)
        for _i in range(0, len(df), _CH):
            _t = pa.Table.from_pandas(df.iloc[_i:_i + _CH], preserve_index=False, schema=_schema, nthreads=1)
            _writer.write_table(_t)
            del _t
            gc.collect()
        _writer.close()
    print(f"\nWrote {out_path} ({len(df):,} rows, {out_path.stat().st_size/1e6:.1f} MB)", flush=True)

    # ----- Summary -----
    n_v = df["vertical_id"].notna().sum()
    n_h = df["horizontal_id"].notna().sum()
    n_v_manual = df["vertical_source"].eq("manual").sum()
    n_h_manual = df["horizontal_source"].eq("manual").sum()
    n_v_rule = df["vertical_source"].astype("string").fillna("").str.startswith("rule").sum()
    n_h_rule = df["horizontal_source"].astype("string").fillna("").str.startswith("rule").sum()
    n_v_acct = df["vertical_source"].astype("string").fillna("").str.startswith("acct_code").sum()
    n_h_acct = df["horizontal_source"].astype("string").fillna("").str.startswith("acct_code").sum()
    n_v_rerank = df["vertical_source"].astype("string").fillna("").str.startswith("rerank").sum()
    n_h_rerank = df["horizontal_source"].astype("string").fillna("").str.startswith("rerank").sum()
    n_v_llm = df["vertical_source"].eq("llm").sum()
    n_h_llm = df["horizontal_source"].eq("llm").sum()

    print(f"\n=== Tag coverage ===")
    print(f"  vertical:    tagged {n_v:,}/{len(df):,}  manual={n_v_manual}  rule={n_v_rule}  rerank={n_v_rerank}  llm={n_v_llm}")
    print(f"  horizontal:  tagged {n_h:,}/{len(df):,}  manual={n_h_manual}  rule={n_h_rule}  acct_code={n_h_acct}  rerank={n_h_rerank}  llm={n_h_llm}")

    print(f"\n=== Top verticals (by row count) ===")
    print(df["vertical_label"].value_counts().head(15).to_string())
    print(f"\n=== Top horizontals (by row count) ===")
    print(df["horizontal_label"].value_counts().head(15).to_string())
    print(f"\n=== Capex / Opex × Gaming / Non-gaming ===")
    print(pd.crosstab(df["final_capex_opex"], df["ng_scope"]).to_string())


if __name__ == "__main__":
    main()
