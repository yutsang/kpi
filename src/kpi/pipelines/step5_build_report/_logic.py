"""Step 5: Build cross-tab report (rows × columns) for client.

Reads: data/interim/<company>_tagged_rows.parquet
Outputs: data/output/<company>_kpi_report.xlsx with multiple sheets:

  - 0_index               : map of all sheets and what they contain
  - 1_master_pivot        : full cross-tab (verticals × horizontals × capex/opex × period × posting_year)
  - 2_by_period           : pivot grouped by period (23/24/25)
  - 3_by_posting_year     : pivot grouped by posting year
  - 4_capex_only          : only capex rows
  - 5_opex_only           : only opex rows
  - 6_audit_drill         : raw row sample for any (vertical, horizontal) cell, for traceability

  - <company>_tagged_rows.parquet is also retained as the PowerBI source — open in PBI
    as a single fact table; build matrix visuals using vertical_label as rows and
    horizontal_label as columns.
"""
from __future__ import annotations

import gc
import sys
from pathlib import Path

import pandas as pd

from kpi.lib.conf import load_config, load_categories, path  # noqa: E402
from kpi.lib.io_setup import force_unbuffered_io  # noqa: E402

force_unbuffered_io()


def _pivot(df: pd.DataFrame, index_cols, column_col, value_col, agg="sum") -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    pv = df.pivot_table(
        index=index_cols,
        columns=column_col,
        values=value_col,
        aggfunc=agg,
        fill_value=0,
        margins=True,
        margins_name="總計",
        observed=False,
    )
    return pv


def _pivot_with_count(
    df: pd.DataFrame,
    index_cols: list[str],
    column_col: str,
    amount_col: str,
    count_by_index: dict | None = None,
) -> pd.DataFrame:
    """Pivot with [數量/次數] | horizontal amount columns | [總計 row sum] + bottom 總計 row.

    If `count_by_index` provided (single-level index only), use it as the H_COUNT
    source instead of row count. Typically passed from step4.5's vertical_counts.parquet.
    """
    if df.empty:
        return pd.DataFrame()
    amt_pv = df.pivot_table(
        index=index_cols,
        columns=column_col,
        values=amount_col,
        aggfunc="sum",
        fill_value=0,
        observed=False,
    )
    if count_by_index is not None and len(index_cols) == 1:
        cnt = pd.Series(
            {k: count_by_index.get(k) for k in amt_pv.index},
            name="數量/次數",
            dtype="Float64",
        )
    else:
        cnt = df.groupby(index_cols, observed=False).size().rename("數量/次數")
    out = pd.concat([cnt, amt_pv], axis=1)
    horizontal_cols = [c for c in out.columns if c != "數量/次數"]
    out["總計"] = out[horizontal_cols].sum(axis=1, numeric_only=True)

    bottom = out.sum(numeric_only=True).to_frame().T
    if isinstance(out.index, pd.MultiIndex):
        bottom.index = pd.MultiIndex.from_tuples(
            [tuple(["總計"] * len(index_cols))], names=index_cols,
        )
    else:
        bottom.index = pd.Index(["總計"], name=index_cols[0])
    return pd.concat([out, bottom])


def main():
    cfg = load_config()
    cats = load_categories()
    company = cfg["company"]["code"]
    interim = path(cfg, "interim_dir")
    out_dir = path(cfg, "output_dir")
    out_dir.mkdir(parents=True, exist_ok=True)
    cols = cfg["columns"]

    src = interim / f"{company}_tagged_rows.parquet"
    if not src.exists():
        raise FileNotFoundError(f"Run step 4 first. Missing: {src}")

    print(f"Loading {src} ...", flush=True)
    df = pd.read_parquet(src)
    print(f"  rows={len(df):,}", flush=True)

    # Normalise column names (strip accidental trailing spaces from Excel)
    df.columns = df.columns.str.strip()
    cols = {k: v.strip() if isinstance(v, str) else v for k, v in cols.items()}

    # Derive period / posting_year from posting_date when not configured
    _date_col = cols.get("posting_date", "")
    if not cols.get("period") or cols["period"] not in df.columns:
        if _date_col and _date_col in df.columns:
            df["_period"] = (pd.to_datetime(df[_date_col], errors="coerce")
                             .dt.to_period("M").astype("string"))
        else:
            df["_period"] = "(未分類)"
        cols = {**cols, "period": "_period"}
    if not cols.get("posting_year") or cols["posting_year"] not in df.columns:
        if _date_col and _date_col in df.columns:
            df["_posting_year"] = (pd.to_datetime(df[_date_col], errors="coerce")
                                   .dt.year.astype("string"))
        elif cols.get("period") and cols["period"] in df.columns:
            # Extract 4-digit year from period strings like "Jan-25", "2025 - Jan", "2025-01"
            _period_s = df[cols["period"]].astype(str)
            _year_full = _period_s.str.extract(r'\b(20\d{2})\b')[0]          # "2025 - Jan"
            _year_short = _period_s.str.extract(r'-(\d{2})$')[0]             # "Jan-25"
            _year_short = _year_short.apply(lambda y: f"20{y}" if pd.notna(y) else pd.NA)
            df["_posting_year"] = _year_full.combine_first(_year_short).fillna("(未分類)")
        else:
            df["_posting_year"] = "(未分類)"
        cols = {**cols, "posting_year": "_posting_year"}

    # ===== Prep =====
    df["amount"] = pd.to_numeric(df[cols["amount"]], errors="coerce").fillna(0)
    df["vertical_label"] = df["vertical_label"].fillna("(未分類)")
    df["horizontal_label"] = df["horizontal_label"].fillna("(未分類)")
    df["final_capex_opex"] = df["final_capex_opex"].fillna("(未分類)")
    df[cols["period"]] = df[cols["period"]].astype("string").fillna("(未分類)")
    df[cols["posting_year"]] = df[cols["posting_year"]].astype("string").fillna("(未分類)")

    # Vertical labels in correct order from categories.yaml (preserve report row order)
    v_order = [v["label"] for v in cats["verticals"]]
    h_order = [h["label"] for h in cats["horizontals"] if h["id"] != "H_COUNT"]

    # Set categorical for ordered display
    df["vertical_label"] = pd.Categorical(df["vertical_label"], categories=v_order + ["(未分類)"], ordered=True)
    df["horizontal_label"] = pd.Categorical(df["horizontal_label"], categories=h_order + ["(未分類)"], ordered=True)

    out_path = out_dir / f"{company}_kpi_report.xlsx"
    out_kpi_only = out_dir / f"{company}_kpi_only.xlsx"
    print(f"Building {out_path} ...", flush=True)

    # Load distinct counts from step 4.5 if available (replaces row-count 數量/次數)
    counts_path = interim / f"{company}_vertical_counts.parquet"
    count_by_label: dict | None = None
    if counts_path.exists():
        vc = pd.read_parquet(counts_path)
        vid_to_label = {v["id"]: v["label"] for v in cats["verticals"]}
        count_by_label = {}
        for _, row in vc.iterrows():
            vid = row["vertical_id"]
            cnt_val = row["count"]
            label = vid_to_label.get(vid)
            if label is not None:
                count_by_label[label] = (
                    int(cnt_val) if pd.notna(cnt_val) else None
                )
        print(f"  Loaded distinct counts from {counts_path.name}  ({len(count_by_label)} verticals)")
    else:
        print(f"  WARNING: {counts_path.name} not found — 數量/次數 will fall back to row count")

    # Master pivot computed once, written to both files.
    master = _pivot_with_count(
        df,
        index_cols=["vertical_label"],
        column_col="horizontal_label",
        amount_col="amount",
        count_by_index=count_by_label,
    )

    # If step0.5 split rows by year-bucket, also build per-bucket pivots.
    has_period = "report_period" in df.columns
    period_pivots: dict[str, pd.DataFrame] = {}
    if has_period:
        # Chronological order: oldest reporting cycle first, current year last.
        period_order = ["23", "24_23SY", "24", "25_23SY", "25_24SY", "25"]
        present = set(df["report_period"].astype(str))
        present_periods = [p for p in period_order if p in present]
        # Defensive: include any other period values found
        present_periods += [p for p in sorted(present) if p not in period_order]
        print(f"  report_period present in tagged_rows — building per-bucket pivots: {present_periods}")

        # Per-period DISTINCT count using same count_cols logic as step4.5 (consistency
        # with master pivot which uses distinct project counts, not row counts).
        raw_cc = cols.get("count_cols")
        if isinstance(raw_cc, str):
            cc_list = [raw_cc]
        elif isinstance(raw_cc, (list, tuple)):
            cc_list = list(raw_cc)
        else:
            cc_list = []
        cc_avail = [c for c in cc_list if c in df.columns]
        vid_to_label = {v["id"]: v["label"] for v in cats["verticals"]}
        for p in present_periods:
            sub = df[df["report_period"].astype(str) == p]
            per_period_counts: dict | None = None
            if cc_avail and "vertical_id" in sub.columns:
                per_period_counts = {}
                for v in cats["verticals"]:
                    v_sub = sub[sub["vertical_id"] == v["id"]]
                    label = v["label"]
                    if len(v_sub) == 0:
                        per_period_counts[label] = 0
                    else:
                        per_period_counts[label] = int(v_sub.groupby(cc_avail, dropna=False).ngroups)
            period_pivots[p] = _pivot_with_count(
                sub, index_cols=["vertical_label"], column_col="horizontal_label",
                amount_col="amount", count_by_index=per_period_counts,
            )

    # DEFAULT: parquet only — no Excel. Set KPI_EXPORT_XLSX=1 to also write kpi_only.xlsx
    # + kpi_report.xlsx pivots (and step6 master-audit 投資方向 xlsx).
    import os
    import contextlib
    _export_xlsx = bool(os.environ.get("KPI_EXPORT_XLSX", "").strip())
    if _export_xlsx:
        with pd.ExcelWriter(out_kpi_only, engine="xlsxwriter") as kpi_writer:
            master.to_excel(kpi_writer, sheet_name="1_master_pivot_all")
            for p, pv in period_pivots.items():
                pv.to_excel(kpi_writer, sheet_name=f"1_master_pivot_{p}")
        print(f"  Wrote {out_kpi_only.name}  ({out_kpi_only.stat().st_size/1e3:.0f} KB)")
    else:
        print("  [step5] kpi_only.xlsx SKIPPED (parquet only). Set KPI_EXPORT_XLSX=1 for Excel.", flush=True)

    with (pd.ExcelWriter(out_path, engine="xlsxwriter") if _export_xlsx else contextlib.nullcontext()) as writer:
        if _export_xlsx:
            # ----- Sheet 1: Master pivot (all periods combined) -----
            master.to_excel(writer, sheet_name="1_master_pivot_all")
            # Per-period pivots (one sheet each)
            for p, pv in period_pivots.items():
                pv.to_excel(writer, sheet_name=f"1_master_pivot_{p}")

        # (top-500 audit sheet removed — the full 大表 / parquet below already hold every row.)

        # ----- Sheet 7: Original JE rows + tags (FULL — no cap, single sheet) -----
        # Auditors / reviewers open ONE file and see the full ledger annotated with
        # vertical / horizontal tags. Sorted by |amount| desc.
        _RESERVED_REPORT_KEYS = {
            "period", "posting_year", "posting_date",
            "amount", "capex_opex", "project", "project_name_cols",
            "ng11_category", "account_code", "account_code_cols",
            "account_desc", "description", "description_cols",
            "vendor", "unique_id", "unique_id_cols", "count_cols",
        }
        base_cols = [
            cols.get("period"), cols.get("posting_year"), cols.get("posting_date"),
            cols.get("project"), cols.get("ng11_category"), cols.get("capex_opex"),
            cols.get("amount"), cols.get("account_code"), cols.get("account_desc"),
            cols.get("description"), cols.get("vendor"),
        ]
        # Auto-append any extra mapped columns (job_code, comp_nature, spend_nature, etc.)
        extra_cols_appended: list[str] = []
        for k, v in cols.items():
            if k in _RESERVED_REPORT_KEYS:
                continue
            if not isinstance(v, str) or not v:
                continue
            if v in df.columns and v not in base_cols:
                extra_cols_appended.append(v)
        tag_cols = [
            "vertical_id", "vertical_label", "vertical_source",
            "horizontal_id", "horizontal_label", "horizontal_source",
            "final_capex_opex", "row_type", "ng_scope",
            # step0.5 split audit columns (present when year_split configured)
            "report_period", "split_ratio", "split_source", "original_amount",
        ]
        # audit_detail_cols raw columns → keep them in kpi_report so the deliverable / Tableau can
        # expose the uniform detail layers (科目層級/科目明細/發票號/PO號/成本中心/WBS子項/憑證號).
        # They live in tagged_rows but aren't in `cols`, so step5 used to trim them away.
        _ent = ((cfg.get("_master") or {}).get("companies") or {}).get(company, {}) or {}
        _adc = cfg.get("audit_detail_cols") or _ent.get("audit_detail_cols") or {}
        _detail_raw = [r for v in _adc.values() for r in (v if isinstance(v, list) else [v]) if r]

        # ── Adjustment dimensions (調整金額 / 調整後金額 / 調整一級 / 調整二級) ────────────────────
        # The project team's adjustment cols are named inconsistently per entity & year: English
        # canon in the unified 23-tie files (adjustment_amount/adjusted_amount/adjust_lv1/lv2), Chinese
        # in the native 24/25 files (調整金額/調整後金額/2024年度調整/…). Coalesce each entity's variants
        # (first non-blank wins) into 4 stable Chinese-named cols so 大表 + Tableau expose them
        # uniformly across ALL years. MGM uses per-bucket col names (25_Adj/24_Adj etc.), handled below.
        _ALIAS = {"company_1": "galaxy", "company_2": "sjm", "company_3": "wynn",
                  "company_4": "vml", "company_5": "melco", "company_6": "mgm"}
        _alias = cfg.get("alias") or _ALIAS.get(company, "")
        # A target maps to either a list of value cols (coalesce first non-blank), or a dict
        # {"cols": [...], "flags": [...]}: after coalescing `cols`, fall back to flag columns —
        # each flag col holds 'Y' when the row IS that adjustment type, so the col's NAME becomes
        # the lv1 value (user-confirmed 24/25 layouts: vml-24 & sjm-25 use the multi-flag form).
        ADJUST_MAP = {
            "galaxy": {"調整金額": ["adjustment_amount", "調整金額"],   # 24/23=adjustment_amount; 25=調整金額(raw)
                       "調整一級": ["adjust_lv1", "一級調整"],   # 25=一級調整(raw); 24/23=adjust_lv1(tied)
                       "調整二級": ["adjust_lv2", "二級調整"]},  # 25=二級調整(raw); 24/23=adjust_lv2(tied)
            "wynn":   {"調整金額": ["adjustment_amount", "調整金額"],
                       "調整一級": ["adjust_lv1", "調整項目名稱-調整數不重合"],
                       "調整二級": ["adjust_lv2", "事項備註"]},
            "vml":    {"調整金額": ["adjustment_amount", "調整金額"],
                       "調整一級": {"cols": ["adjust_lv1"], "flags": [
                           "商業性會展項目支出", "其他日常營運支出",
                           "與投資項目無關的贈房支出、免費餐飲支出及贈票支出",
                           "不符合“吸引外國客源”定義的相關投資支出",
                           "未在原投資計劃明確列示且其後未被認可的新增項目投資支出",
                           "未完全實現投資目的的投資支出",
                           "計入報告投資金額的2025年投資計劃獲批後的支出",
                           "編制報告投資金額過程中的取數差錯", "不被認可的支出"]},
                       "調整二級": ["adjust_lv2"]},
            "melco":  {"調整金額": ["adjustment_amount", "調整金額"],
                       "調整一級": ["adjust_lv1", "初步識別調整類型", "調整類別"],
                       "調整二級": ["adjust_lv2", "初步識別調整事項", "調整事項"]},
            "sjm":    {"調整金額": ["調整金額", "2024年度調整", "2023年度調整"],
                       "調整一級": {"cols": ["adjust_lv1"], "flags": [
                           "用於計算酒店贈房支出的贈房單價超過ADR部分的支出", "酒店客房改造支出",
                           "不符合“吸引外國客源”定義的相關投資支出",
                           "未在原投資計劃明確列示且其後未被認可的新增項目投資支出",
                           "2026年1-4月入賬的支出計入2025年報告投資金額",
                           "不符合期後事項定義的投資金額",
                           "未完全實現投資目的的投資支出"]},
                       "調整二級": ["adjust_lv2", "調整事項"]},
            "mgm":    {"調整一級": ["Adj_lv1", "adjust_lv1", "調整項目名稱"],
                       "調整二級": ["Adj_lv2", "adjust_lv2", "調整事項備註", "調整事項"]},
        }
        _BLNK = ["", "nan", "None", "NaN", "<NA>", "NaT"]
        _NUM_TGT = {"調整金額", "調整後金額"}   # amounts → numeric (Tableau '#' measure); lv1/lv2 stay text
        _adjust_names: list[str] = []
        for _tgt, _spec in (ADJUST_MAP.get(_alias) or {}).items():
            _cols = _spec.get("cols", []) if isinstance(_spec, dict) else _spec
            _flags = _spec.get("flags", []) if isinstance(_spec, dict) else []
            if not any(c in df.columns for c in _cols) and not any(f in df.columns for f in _flags):
                continue
            _out = pd.Series("", index=df.index, dtype="object")
            for _c in _cols:
                if _c in df.columns:
                    _s = df[_c].astype("string").fillna("").str.strip()
                    _out = _out.mask(_out.isin(_BLNK), _s)
            for _f in _flags:
                if _f in df.columns:
                    _isY = df[_f].astype("string").fillna("").str.strip().str.upper().eq("Y")
                    _out = _out.mask(_out.isin(_BLNK) & _isY, _f)
            if _tgt in _NUM_TGT:
                # numeric so Tableau treats it as a measure (#), not text (Abc); blank → NaN
                df[_tgt] = pd.to_numeric(_out.astype(str).str.replace(",", "", regex=False),
                                         errors="coerce")
            else:
                df[_tgt] = _out.replace({"nan": "", "None": ""})
            _adjust_names.append(_tgt)
        if _adjust_names:
            print(f"  [adjust] coalesced {_adjust_names} for {_alias}", flush=True)
        # 調整後金額: KEEP the project team's native post-adjustment col where a row has one, ELSE COMPUTE
        # = amount + 調整金額 (added on 25-buckets only — 24/23 amount is already post-adjustment, so
        # adding would double-count). Native per entity: galaxy adjusted_amount / vml 調整後金額 /
        # sjm 2023年度調整后. wynn/melco: fully computed. MGM: per-bucket (25_Adj/24_Adj/24_調整金額/etc.)
        # handled below and stored in df["調整後金額"] before _NATIVE_POST picks it up.
        _amt_col = cols.get("amount")
        if _amt_col and _amt_col in df.columns:
            # ── MGM per-bucket 調整金額 & 調整後金額 ──────────────────────────────────────────────
            # 25-file: 25_Adj/24_Adj (no native post-adj); 24-file: 24_調整金額/24_調整后/23_調整金額/
            # 23_調整后; 23-file: 調整金額/調整後金額. Cross-bucket coalesce is unsafe (wide-format rows
            # carry all year cols even after melt), so resolve by report_period prefix.
            if _alias == "mgm" and "report_period" in df.columns:
                _rp_m = df["report_period"].astype(str).fillna("")
                _adj_m = pd.Series(float("nan"), index=df.index, dtype="float64")
                for _pfx, _c in [("25", "25_Adj"), ("24_23", "23_調整金額"),
                                  ("24", "24_Adj"), ("24", "24_調整金額"), ("23", "調整金額")]:
                    if _c not in df.columns: continue
                    _mask = _rp_m.str.startswith(_pfx) & _adj_m.isna()
                    _adj_m = _adj_m.where(~_mask, pd.to_numeric(df[_c], errors="coerce"))
                df["調整金額"] = _adj_m
                if "調整金額" not in _adjust_names: _adjust_names.append("調整金額")
                _post_m = pd.Series(float("nan"), index=df.index, dtype="float64")
                for _pfx, _c in [("24_23", "23_調整后"), ("24", "24_調整后"), ("23", "調整後金額")]:
                    if _c not in df.columns: continue
                    _mask = _rp_m.str.startswith(_pfx) & _post_m.isna()
                    _post_m = _post_m.where(~_mask, pd.to_numeric(df[_c], errors="coerce"))
                _base_m = pd.to_numeric(df[_amt_col], errors="coerce")
                df["調整後金額"] = _post_m.where(_post_m.notna(), _base_m + _adj_m.fillna(0.0))
                if "調整後金額" not in _adjust_names: _adjust_names.append("調整後金額")
                print(f"  [adjust-mgm] per-bucket 調整金額/後 wired", flush=True)
            _NATIVE_POST = {"galaxy": ["adjusted_amount", "調整後金額"],   # 24/23=adjusted_amount; 25=調整後金額(raw)
                            "vml": ["調整後金額"], "sjm": ["2023年度調整后"],
                            "mgm": ["調整後金額"]}   # use MGM's pre-computed per-bucket value above
            _native = pd.Series("", index=df.index, dtype="object")
            for _c in _NATIVE_POST.get(_alias, []):
                if _c in df.columns:
                    _s = df[_c].astype("string").fillna("").str.strip()
                    _native = _native.mask(_native.isin(_BLNK), _s)
            _native_num = pd.to_numeric(_native.astype(str).str.replace(",", "", regex=False),
                                        errors="coerce")
            _base = pd.to_numeric(df[_amt_col], errors="coerce")
            _delta = (pd.to_numeric(df["調整金額"], errors="coerce").fillna(0.0)
                      if "調整金額" in df.columns else pd.Series(0.0, index=df.index))
            _rp = (df["report_period"].astype("string").fillna("")
                   if "report_period" in df.columns else pd.Series("", index=df.index))
            _computed = _base + _delta.where(_rp.str.startswith("25"), 0.0)
            df["調整後金額"] = _native_num.where(_native_num.notna(), _computed)
            if "調整後金額" not in _adjust_names:
                _adjust_names.append("調整後金額")
            print(f"  [adjust] 調整後金額 = native(if any) else amount+調整金額(25) for {_alias}", flush=True)
        # unified-raw reference cols (項目組 own labels — reference only, OUR taxonomy is canonical)
        # unified-raw extra cols — carry through so the 大表 / Tableau keep EVERY column the
        # project team put in the tied raw (user 2026-06-14: "take care of any column 包括 remarks").
        _UNIFIED_EXTRA = ["unique_id", "entity", "year", "project_code", "dicj_code", "subproject",
                          "ng_theme", "amount_mop", "adjustment_amount", "adjusted_amount",
                          "adjust_lv1", "adjust_lv2", "pt_class_H", "pt_class_V", "source",
                          "comp_type", "is_labor", "is_internal", "take_flag", "take_flag2",
                          "netoff_flag", "internal", "remark", "25跨年", "24跨年", "23跨年"]
        _ref_cols = [c for c in (_UNIFIED_EXTRA + _adjust_names) if c in df.columns and c not in base_cols]
        # 4 層 project 結構嘅原生細碼/細名欄 — 保證 survive 落 parquet 畀 prep_tableau wire（SUBPROJECT_COLS）。
        # 只 keep 存在嘅，對冇呢啲欄嘅 entity 無害。
        _PROJECT_STRUCT = ["Project & Sub-project ID", "Project name - Amended", "project_mre_id",
                           "Project_name", "Project/Sub-Project Name_1", "Sub-Project Serial Number",
                           "Subproject", "SubProject_Name", "Sub project", "项目名称中文"]
        _proj_cols = [c for c in _PROJECT_STRUCT if c in df.columns and c not in base_cols]
        all_row_cols = base_cols + extra_cols_appended + _detail_raw + _ref_cols + _proj_cols + tag_cols
        all_row_cols = [c for c in all_row_cols if c and c in df.columns]
        # Dedupe (preserve order) — pandas to_parquet fails on duplicate col names
        seen = set()
        all_row_cols = [c for c in all_row_cols if not (c in seen or seen.add(c))]
        # Stream the full dataset to parquet WITHOUT ever building a second full-frame object.
        # The old `df.loc[order, all_row_cols]` reorder copied EVERY column at once and OOM'd even
        # on a 1 MiB int8 array — that is an ADDRESS-SPACE ceiling (32-bit Python can only map
        # ~2 GB no matter how much RAM the box has), not free-RAM exhaustion. By slicing row-chunks
        # straight from df, peak memory = df + one small chunk (no duplicate frame). Parquet row
        # order is irrelevant for a fact table; downstream sorts as needed.
        keep = [c for c in all_row_cols if c in df.columns]
        _seen = set()
        keep = [c for c in keep if not (c in _seen or _seen.add(c))]
        grand_total = float(df["amount"].sum())
        n_rows = len(df)

        parquet_path = out_path.with_suffix(".parquet")
        import pyarrow as pa
        import pyarrow.parquet as pq
        try:
            # Direct write first — infers the schema from the WHOLE frame (so a column that is
            # all-null in the first rows but has values later, e.g. Melco 人工成本分類, types
            # correctly). Works for small/medium entities; only Galaxy (1.15M) needs streaming.
            df[keep].to_parquet(parquet_path, index=False)
            print(f"  Wrote {parquet_path.name} ({parquet_path.stat().st_size/1e6:.1f} MB, "
                  f"{n_rows:,} rows — full, untruncated)", flush=True)
        except Exception as _e:
            print(f"  [warn] direct parquet failed ({type(_e).__name__}: {_e}); chunked stream…", flush=True)
            _CH = 200_000
            _schema = pa.Table.from_pandas(df.iloc[:1000][keep], preserve_index=False, nthreads=1).schema
            # A column all-null in the 1000-row sample infers as null type → force string so later
            # non-null chunks convert instead of raising 'Invalid null value'.
            _schema = pa.schema([f if not pa.types.is_null(f.type) else pa.field(f.name, pa.string())
                                 for f in _schema])
            _w = pq.ParquetWriter(parquet_path, _schema)
            for _i in range(0, n_rows, _CH):
                _chunk = df.iloc[_i:_i + _CH][keep]
                _w.write_table(pa.Table.from_pandas(_chunk, preserve_index=False, schema=_schema, nthreads=1))
                del _chunk
                gc.collect()
            _w.close()
            print(f"  Wrote {parquet_path.name} ({parquet_path.stat().st_size/1e6:.1f} MB, "
                  f"{n_rows:,} rows — streamed)", flush=True)
        del df
        gc.collect()

        # DEFAULT: skip per-year all_rows xlsx (heavy I/O during iteration).
        # parquet (above) is always written and is the canonical source.
        # For delivery, set KPI_GENERATE_ALL_ROWS_XLSX=1 to generate xlsx files.
        if not os.environ.get("KPI_GENERATE_ALL_ROWS_XLSX", "").strip():
            print(f"  [step5] all_rows xlsx SKIPPED (parquet only). Set KPI_GENERATE_ALL_ROWS_XLSX=1 for delivery.", flush=True)
            return
        # delivery path needs the frame back — read it from the parquet we just streamed
        df_all = pd.read_parquet(parquet_path)

        # Also write per-year all_rows xlsx files (split by leading 2 chars of
        # report_period: "23", "24" incl "24_2", "25" incl "25_2"). Each file
        # stays well under Excel's 1,048,575 row limit. Files are separate from
        # the main kpi_report.xlsx to keep that file small for clients.
        df_xlsx = df_all[all_row_cols].copy()
        if "report_period" in df_xlsx.columns:
            df_xlsx["_year_bucket"] = df_xlsx["report_period"].astype(str).str[:2]
            year_buckets = sorted(df_xlsx["_year_bucket"].unique())
        else:
            df_xlsx["_year_bucket"] = "all"
            year_buckets = ["all"]

        XLSX_LIMIT = 1_048_574  # one less than the actual limit to leave room for header
        for yr in year_buckets:
            sub = df_xlsx[df_xlsx["_year_bucket"] == yr].drop(columns=["_year_bucket"])
            if len(sub) == 0:
                continue
            year_path = out_path.with_name(f"{out_path.stem}_{yr}_all_rows.xlsx")
            try:
                if len(sub) <= XLSX_LIMIT:
                    sub.to_excel(year_path, index=False, sheet_name=f"all_rows_{yr}",
                                 engine="xlsxwriter")
                    print(f"  Wrote {year_path.name} ({year_path.stat().st_size/1e6:.1f} MB, "
                          f"{len(sub):,} rows)", flush=True)
                else:
                    # Even single-year exceeds limit — split into chunks
                    n_chunks = (len(sub) + XLSX_LIMIT - 1) // XLSX_LIMIT
                    for i in range(n_chunks):
                        chunk = sub.iloc[i * XLSX_LIMIT:(i + 1) * XLSX_LIMIT]
                        chunk_path = out_path.with_name(
                            f"{out_path.stem}_{yr}_all_rows_part{i + 1}of{n_chunks}.xlsx"
                        )
                        chunk.to_excel(chunk_path, index=False,
                                       sheet_name=f"all_rows_{yr}_p{i + 1}",
                                       engine="xlsxwriter")
                        print(f"  Wrote {chunk_path.name} "
                              f"({chunk_path.stat().st_size/1e6:.1f} MB, {len(chunk):,} rows)",
                              flush=True)
            except Exception as e:
                print(f"  [warn] xlsx write failed for {yr}: {e}", flush=True)

    if out_path.exists():
        print(f"\nWrote {out_path}  ({out_path.stat().st_size/1e6:.1f} MB)")
    print(f"  Total amount: {grand_total:,.0f}  rows: {n_rows:,}")
    print(f"\nFor PowerBI: use {src.name} as fact table.")
    print(f"  Matrix visual: rows=vertical_label, columns=horizontal_label, values=Sum(Reported Amount(MOP))")
    print(f"  Slicers: final_capex_opex, Period, Posting Year, ng_scope")


if __name__ == "__main__":
    main()
