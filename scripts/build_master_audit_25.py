"""Build per-entity master audit Excel — FAST version, 24 + 25.

Drops slow per-project pivots. Each entity gets ONE Excel with sheets:

  0_index             — what each sheet contains
  1_KPI_25            — V × H cross-tab matrix for year 25, Σ amount_mop
  2_KPI_24            — V × H cross-tab matrix for year 24
  3_category_ref      — V + H category reference (id, label, NG mapping)
Output: data/{ent}/output/{ent}_master_audit.xlsx  (~10-30s per entity)

Run:
  python scripts/build_master_audit_25.py
  python scripts/build_master_audit_25.py --entity galaxy
"""
from __future__ import annotations
import argparse, gc, re, sys, time
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
import pandas as pd
import yaml

ENTITIES = {"galaxy":"company_1","sjm":"company_2","wynn":"company_3",
            "vml":"company_4","melco":"company_5","mgm":"company_6"}

V_TO_NG = {
    "V_GAMING_VENUE": ("NG0", "博彩項目"),
    "V_GAMING_EQUIP": ("NG0", "博彩項目"),
    "V_OVERSEAS_OFFICE": ("NG1", "吸引外國客源"),
    "V_OVERSEAS_WEB_SEO": ("NG1", "吸引外國客源"),
    "V_OVERSEAS_ROADSHOW": ("NG1", "吸引外國客源"),
    "V_INVITE_GUEST": ("NG1", "吸引外國客源"),
    "V_INVITE_AGENCY": ("NG1", "吸引外國客源"),
    "V_REGIONAL_TEAM": ("NG1", "吸引外國客源"),
    "V_REGIONAL_SALES": ("NG1", "吸引外國客源"),
    "V_PROMO_VIDEO": ("NG1", "吸引外國客源"),
    "V_MICE": ("NG2", "會議展覽"),
    "V_CONCERT": ("NG3", "娛樂表演"),
    "V_SPORT_EVENT": ("NG4", "體育盛事"),
    "V_VENUE_PERF_SPORT_MICE": ("NG4", "體育盛事"),
    "V_ART_EXHIBITION": ("NG5", "文化藝術"),
    "V_MUSEUM": ("NG5", "文化藝術"),
    "V_WELLNESS": ("NG6", "健康養生"),
    "V_THEME_PARK": ("NG7", "主題遊樂"),
    "V_RESTAURANT": ("NG8", "美食之都"),
    "V_FOOD_EVENT": ("NG8", "美食之都"),
    "V_COMMUNITY": ("NG9", "社區旅遊"),
    "V_MARITIME": ("NG10", "海上旅遊"),
    # V label ≠ NG (not 1:1) — neutral fallback for entities with no raw NG column.
    # 內部設施升級 already exists as V_PROPERTY_UPGRADE (not re-added). 節日慶典 / 公共設施設備提升
    # are 0-row placeholders until the project team picks the projects.
    "V_FESTIVAL": ("NG11", "其他"),             # 節日慶典
    "V_PUBLIC_FACILITY": ("NG11", "其他"),      # 公共設施設備提升
    "V_PROPERTY_UPGRADE": ("NG11", "其他"),
    "V_OTHER": ("NG11", "其他"),
}

SUBPROJECT_CANDIDATES = (
    "Sub project", "SubProject_Name", "Subproject_Name", "subproject",
    "Subproject",
    "项目名称中文", "项目英文名称", "Initiative Name", "Contents Name",
    "initiative Name", "initiative_name",
    "Project/Sub-Project Name", "Project code - Amended",
    "项目名称&Unique ID", "项目名稱",
    "subproject_name", "sub_project", "sub_project_name",
)
YEAR_CANDIDATES = ("report_period", "report_year", "Yr related", "years")


def _cn_kw(s) -> str:
    """Robust 中文 項目性質 → NG (keyword; label variants). '' if none. 博彩 before 娛樂; 非博彩 = noise."""
    s = str(s)
    if "非博彩" in s:
        return ""
    for kws, ng in [(["博彩", "gaming"], "NG0"), (["海上"], "NG10"),
                    (["外國", "客源", "國際客"], "NG1"), (["會議", "會展", "mice"], "NG2"),
                    (["娛樂", "演唱", "表演"], "NG3"), (["體育", "賽事"], "NG4"),
                    (["文化", "藝術", "文藝"], "NG5"), (["健康", "養生"], "NG6"),
                    (["主題", "遊樂"], "NG7"), (["美食", "餐飲"], "NG8"),
                    (["社區"], "NG9"), (["其他"], "NG11")]:
        if any(k in s or k in s.lower() for k in kws):
            return ng
    return ""


def _fuzzy_col(df, name):
    if not name: return None
    if name in df.columns: return name
    for c in df.columns:
        if str(c).strip() == str(name).strip(): return c
    return None


def _filter_year(df, ycol, year_prefix):
    """Keep rows whose year column starts with year_prefix (e.g. '25' or '24')."""
    # .astype("string") (pandas StringDtype) — categorical-safe, never numpy fixed-width <U;
    # fillna("") so the boolean mask has no NA (NA in a mask raises on df[mask]).
    s = df[ycol].astype("string").fillna("")
    return df[s.str.startswith(year_prefix) | (s == f"Yr 20{year_prefix}") | (s == f"20{year_prefix}")].copy()


def _kpi_pivot(df, categories=None):
    """V × H cross-tab with NG header, Σ amount_mop + grand totals.

    Reindexes the H columns to the FULL categories.yml horizontal list so a category that
    happens to be empty for this entity (e.g. SJM has no 活動場地 / 合約成本演藝) still shows
    as a 0 column instead of vanishing. Any H present in data but not in the taxonomy
    (e.g. a not-yet-defined id) is appended, never dropped."""
    if df.empty or "vertical_label" not in df.columns or "horizontal_label" not in df.columns:
        return pd.DataFrame()
    # Force string dtype on group keys to avoid Cartesian-product OOM from
    # categorical levels carried over from upstream parquet (Galaxy hit this).
    pv_df = df[["ng_code", "ng_label", "vertical_id", "vertical_label",
                "horizontal_label", "amount_mop"]].copy()
    for c in ("ng_code", "ng_label", "vertical_id", "vertical_label", "horizontal_label"):
        # astype(object), NOT astype(str): on a categorical column .astype(str) builds a numpy
        # fixed-width <U{maxlen} array (every row padded to the longest string × 4 bytes →
        # hundreds of MB per column, Galaxy step6 OOM'd at 456 MiB). object keeps cheap pointers.
        pv_df[c] = pv_df[c].astype(object)
    pv = pv_df.pivot_table(
        index=["ng_code", "ng_label", "vertical_id", "vertical_label"],
        columns="horizontal_label",
        values="amount_mop",
        aggfunc="sum", fill_value=0,
        margins=True, margins_name="總計",
        observed=True,
    )
    if categories:
        # H columns in canonical categories.yml order (same as the kpi report 橫向 order)
        h_order = [h["label"] for h in categories.get("horizontals", [])
                   if h.get("id") != "H_COUNT" and h.get("label")]
        extra = [c for c in pv.columns if c not in h_order and c != "總計"]
        cols = h_order + extra + (["總計"] if "總計" in pv.columns else [])
        pv = pv.reindex(columns=cols, fill_value=0)
        # Sort rows by NG (NG0→NG11, 總計 last), then canonical vertical order within each NG.
        v_order = {v["id"]: i for i, v in enumerate(categories.get("verticals", []))}

        def _ngkey(t):
            if str(t[0]) == "總計" or str(t[2]) == "總計":
                return (9999, 9999)
            m = re.match(r"NG(\d+)", str(t[0]))
            return (int(m.group(1)) if m else 998, v_order.get(str(t[2]), 998))
        pv = pv.iloc[sorted(range(len(pv)), key=lambda i: _ngkey(pv.index[i]))]
    return pv


def _category_ref(categories: dict) -> pd.DataFrame:
    """Reference table: V_id, V_label, NG_code, NG_label for verticals;
    H_id, H_label for horizontals. Two sections in one sheet."""
    v_rows = []
    for v in categories.get("verticals", []):
        vid = v.get("id"); vlabel = v.get("label")
        ng_code, ng_label = V_TO_NG.get(vid, ("NG11", "其他"))
        v_rows.append({"category": "V", "id": vid, "label": vlabel,
                       "ng_code": ng_code, "ng_label": ng_label})
    h_rows = []
    for h in categories.get("horizontals", []):
        h_rows.append({"category": "H", "id": h.get("id"), "label": h.get("label"),
                       "ng_code": "", "ng_label": ""})
    return pd.DataFrame(v_rows + h_rows)


def build(ent: str, com: str, categories: dict) -> Path | None:
    t0 = time.time()
    parquet = Path(f"data/{ent}/output/{com}_kpi_report.parquet")
    if not parquet.exists():
        print(f"[{ent}] {parquet} missing — skip", flush=True); return None

    cfg = yaml.safe_load(Path(f"conf/{com}/parameters.yml").read_text(encoding="utf-8"))
    cols = cfg.get("columns", {}) or {}

    df = pd.read_parquet(parquet)
    print(f"[{ent}] loaded {len(df):,} rows in {time.time()-t0:.1f}s", flush=True)

    # 維護費 capex → 建設與設施支出 (same rule as step4; applied here too so it shows on a step6-only
    # re-run without re-running the whole pipeline).
    if "horizontal_id" in df.columns and "final_capex_opex" in df.columns:
        _mc = df["horizontal_id"].astype(str).eq("H_MAINTENANCE") & df["final_capex_opex"].astype(str).eq("Capex")
        if _mc.any():
            df.loc[_mc, "horizontal_id"] = "H_CONSTRUCTION"
            print(f"[{ent}] 維護費(capex) → 建設與設施支出: {int(_mc.sum()):,} rows", flush=True)
    # Refresh V/H display labels from the CURRENT categories.yml so taxonomy RENAMES show in the
    # audit with a step6-only re-run (IDs are source of truth; labels are display-only).
    _vlab = {v["id"]: v.get("label", "") for v in categories.get("verticals", [])}
    _hlab = {h["id"]: h.get("label", "") for h in categories.get("horizontals", [])}
    if "vertical_id" in df.columns:
        _vl = df["vertical_id"].astype(str).map(_vlab)
        df["vertical_label"] = _vl.where(_vl.notna() & _vl.ne(""), df.get("vertical_label", ""))
    if "horizontal_id" in df.columns:
        _hl = df["horizontal_id"].astype(str).map(_hlab)
        df["horizontal_label"] = _hl.where(_hl.notna() & _hl.ne(""), df.get("horizontal_label", ""))

    # Normalize columns
    amt_col = _fuzzy_col(df, cols.get("amount", ""))
    if amt_col:
        df["amount_mop"] = pd.to_numeric(df[amt_col], errors="coerce").fillna(0)
    else:
        df["amount_mop"] = 0.0

    # NG NEVER from V (V_TO_NG) — only from the databook col below. Init blank → unmapped = (未分類).
    df["ng_code"] = ""; df["ng_label"] = ""

    # NG ALWAYS from the databook 範疇 column — NEVER V. Multi-year files may NAME it differently
    # (SJM: 25 '項目類型', 24 '項目性質(e.g.博彩娛樂)') → COALESCE top-level + yearly_sources overrides.
    from kpi.lib.conf import load_categories as _lc
    from kpi.pipelines.step2_tag_projects._logic import normalize_ng_code as _nz
    _cats = _lc()

    def _resolve(x):
        for cand in (x, x.upper().replace(" ", "")):
            r = _nz(cand, _cats) or ""
            if r[:2] == "NG" and r[2:].isdigit():
                return r
        return _cn_kw(x)
    _ng_cols = []
    _conf_names = [cols.get("ng11_category", "")] + [
        (ys.get("columns_override") or {}).get("ng11_category") for ys in (cfg.get("yearly_sources") or [])]
    for _nm in _conf_names:
        _fc = _fuzzy_col(df, _nm)
        if _fc and _fc not in _ng_cols:
            _ng_cols.append(_fc)
    for _c in df.columns:
        if _c not in _ng_cols and any(k in str(_c) for k in ("項目性質", "項目類型", "項目分類", "範疇", "NG11 Category", "NG Category")):
            _ng_cols.append(_c)
    _nglab = {ng: lbl for ng, lbl in V_TO_NG.values()}
    _ngc_series = pd.Series("", index=df.index, dtype="object")
    for _fc in _ng_cols:
        _m = {x: _resolve(x) for x in set(df[_fc].astype(str).unique())}
        _r = df[_fc].astype(str).map(_m).fillna("")
        _r = _r.where(_r.str.fullmatch(r"NG\d+").fillna(False), "")
        _ngc_series = _ngc_series.mask(_ngc_series.eq(""), _r)
    df["ng_code"] = _ngc_series
    df["ng_label"] = df["ng_code"].map(lambda n: _nglab.get(str(n), "(未分類)"))
    _nm = int(_ngc_series.str.fullmatch(r"NG\d+").fillna(False).sum())
    print(f"[{ent}] NG from databook col(s) {_ng_cols}: {_nm:,}/{len(df):,} mapped ({len(df) - _nm:,} → 未分類)", flush=True)

    ycol = next((c for c in YEAR_CANDIDATES if c in df.columns), None)
    if ycol:
        df["year_bucket"] = df[ycol].astype(str)
    else:
        df["year_bucket"] = "?"

    proj_col = _fuzzy_col(df, cols.get("project", "")) or next(
        (c for c in ("Project","Name of Investment Project","SubProject_Name","Sub project")
         if c in df.columns), None)
    # Subproject candidate — but skip if same col as project (avoid duplicate)
    sub_col = next((c for c in SUBPROJECT_CANDIDATES
                    if c in df.columns and c != proj_col), None)
    # Try project_name_cols from conf as secondary subproject sources
    if not sub_col:
        for c in cols.get("project_name_cols", []) or []:
            if c in df.columns and c != proj_col:
                sub_col = c; break
    ac_col = _fuzzy_col(df, cols.get("account_code", ""))
    ad_col = _fuzzy_col(df, cols.get("account_desc", ""))
    dn_col = _fuzzy_col(df, cols.get("description", ""))
    vd_col = _fuzzy_col(df, cols.get("vendor", ""))
    ui_col = _fuzzy_col(df, cols.get("unique_id", "")) or next(
        (c for c in ("唯一識別碼", "KPMG唯一識別碼", "唯一碼", "Unique ID", "unique_id") if c in df.columns), None)
    for src, tgt in [(proj_col, "project"), (sub_col, "subproject"),
                     (ac_col, "account_code"), (ad_col, "account_desc"),
                     (dn_col, "description"), (vd_col, "vendor"), (ui_col, "unique_id")]:
        if src and src in df.columns:
            df[tgt] = df[src].astype(str).str.strip().str.replace(r"[\r\n]+", " ", regex=True)
    # Ensure consistent schema across all 12 files — empty col if entity has no source
    for tgt in ("project", "subproject", "account_code", "account_desc", "description", "vendor", "unique_id"):
        if tgt not in df.columns:
            df[tgt] = ""

    # ── SJM only: internal-resource (admin) comp ─────────────────────────────────
    # subproject = WBS element + CO object name; drop 是否計入內部資源=Y from main and
    # re-inject the comp from Admin Comp summary v2 (BKD + combined) — runs INSIDE the
    # pipeline (step6 / generate) so it's never forgotten. Injected rows tagged year 25.
    if com == "company_2":
        _wbs = next((c for c in df.columns if "WBS element" in str(c)), None)
        _co = next((c for c in df.columns if "CO object" in str(c)), None)
        if _wbs and _co:
            df["subproject"] = (df[_wbs].astype(str).str.strip() + " | " + df[_co].astype(str).str.strip()).str.strip(" |")
        # 表1 for GUEST_FULL_NAME → Description → V match — capture BEFORE dropping AN=Y so
        # comp guests described on internal-comp rows still resolve their project's V.
        _je_full = df
        _je_desc = next((c for c in df.columns if str(c).strip().lower() == "description"), None) \
            or next((c for c in df.columns if "description" in str(c).lower() or "摘要" in str(c)), None) \
            or "description"
        # internal-comp (是否計入內部資源=Y) is excluded at INGESTION via conf exclude_where (step0),
        # same as Net-off — so kpi_report.parquet already omits it and the summary comp below is NOT
        # double-counted. Backstop: if the Y/N flag column still survives into the parquet, drop Y here.
        _an = next((c for c in df.columns
                    if ("內部資源" in str(c) or "計入" in str(c))
                    and df[c].astype(str).str.strip().str.upper().isin(["Y", "N", "", "NAN"]).mean() > 0.8), None)
        if _an:
            _before = len(df)
            df = df[~df[_an].astype(str).str.strip().str.upper().eq("Y")].copy()
            if _before != len(df):
                print(f"[sjm] backstop dropped {_before - len(df):,} residual internal-comp rows ({_an}=Y)", flush=True)
        try:
            import inject_sjm_admin_comp as _inj
            _acrows, _ = _inj.build_rows("data/sjm/raw/Admin Comp summary v2.xlsx",
                                         je_df=_je_full, je_desc_col=_je_desc, je_vid_col="vertical_id")
            if _acrows:
                # aggregate guest/txn-level comp -> (V, H, subproject, source) so the audit 大表
                # stays readable (pivot totals are identical either way)
                _agg = {}
                for _r in _acrows:
                    _k = (_r[0], _r[1], str(_r[2]), _r[4])
                    _agg[_k] = _agg.get(_k, 0.0) + float(_r[3])
                _vlab = {v["id"]: v.get("label", v["id"]) for v in categories.get("verticals", [])}
                _hlab = {h["id"]: h.get("label", h["id"]) for h in categories.get("horizontals", [])}
                _syn = pd.DataFrame([{
                    "entity": ent, "year_bucket": "25", "amount_mop": _amt,
                    "vertical_id": _v, "vertical_label": _vlab.get(_v, _v),
                    "horizontal_id": _h, "horizontal_label": _hlab.get(_h, _h),
                    "ng_code": V_TO_NG.get(str(_v), ("NG11", "其他"))[0],
                    "ng_label": V_TO_NG.get(str(_v), ("NG11", "其他"))[1],
                    "ng_scope": "non_gaming", "final_capex_opex": "", "row_type": f"admin_comp:{_src}",
                    "unique_id": "(admin comp)", "project": "(admin comp)", "subproject": str(_sub),
                    "account_code": "", "account_desc": "admin comp", "description": str(_src), "vendor": "",
                } for (_v, _h, _sub, _src), _amt in _agg.items()])
                # CRITICAL: _filter_year keys off ycol (report_period), NOT year_bucket — stamp it
                # or these synthetic rows get NaN there and are dropped from BOTH year files.
                if ycol:
                    _syn[ycol] = "25"
                df = pd.concat([df, _syn], ignore_index=True)
                print(f"[sjm] admin-comp injected: {len(_syn):,} agg rows ({len(_acrows):,} raw txns), "
                      f"{_syn['amount_mop'].sum():,.0f} MOP -> year 25", flush=True)
        except Exception as _e:
            print(f"[sjm] admin-comp inject skipped: {_e}", flush=True)

    out_dir = Path("data/review")
    out_dir.mkdir(parents=True, exist_ok=True)
    written = []
    XLSX_LIMIT = 1_048_574  # Excel max rows per sheet

    # Flat JE-row column order
    JE_KEEP = ["entity", "year_bucket", "amount_mop",
               "ng_code", "ng_label", "ng_scope",
               "vertical_id", "vertical_label",
               "horizontal_id", "horizontal_label",
               "final_capex_opex", "row_type",
               "unique_id", "project", "subproject",
               "account_code", "account_desc", "description", "vendor"]

    df["entity"] = ent
    if "ng_scope" not in df.columns:
        df["ng_scope"] = ""

    # Shrink the resident frame BEFORE the per-year subset copies. Galaxy (1.15M rows) OOM'd at
    # _filter_year on year 24 — a 4 MiB column copy failed because the box was at its commit limit.
    # Categorizing the low-cardinality columns means a year-subset copies only int codes (not the
    # string payloads), so the copy becomes cheap. description / unique_id stay object (near-unique).
    _CAT = ["entity", "year_bucket", "ng_code", "ng_label", "ng_scope",
            "vertical_id", "vertical_label", "horizontal_id", "horizontal_label",
            "final_capex_opex", "row_type", "project", "subproject",
            "account_code", "account_desc", "vendor"]
    if ycol:
        _CAT.append(ycol)
    for _c in _CAT:
        if _c in df.columns and str(df[_c].dtype) == "object":
            try:
                df[_c] = df[_c].astype("category")
            except Exception:
                pass
    gc.collect()

    def _horizontal_breakdown(sub):
        """橫向 (H-axis drill): each row = (H, account_code, account_desc, project, subproject) sum amount."""
        cols = [c for c in ("horizontal_id", "horizontal_label",
                            "account_code", "account_desc",
                            "project", "subproject") if c in sub.columns]
        if not cols or "amount_mop" not in sub.columns:
            return pd.DataFrame()
        # Force string dtype to avoid Cartesian-product OOM on categorical levels.
        gdf = sub[cols + ["amount_mop"]].copy()
        for c in cols:
            gdf[c] = gdf[c].astype(object)   # object not str — str on a categorical → numpy <U OOM
        g = gdf.groupby(cols, observed=True, dropna=False)["amount_mop"].agg(["sum", "size"]).reset_index()
        g = g.rename(columns={"sum": "amount_mop", "size": "筆數"})
        return g.sort_values(["horizontal_label", "amount_mop"] if "horizontal_label" in g.columns
                             else ["amount_mop"], ascending=[True, False])

    def _vertical_breakdown(sub):
        """縱向 (V-axis drill): each row = (NG, V, project, subproject) sum amount."""
        cols = [c for c in ("ng_code", "ng_label",
                            "vertical_id", "vertical_label",
                            "project", "subproject") if c in sub.columns]
        if not cols or "amount_mop" not in sub.columns:
            return pd.DataFrame()
        gdf = sub[cols + ["amount_mop"]].copy()
        for c in cols:
            gdf[c] = gdf[c].astype(object)   # object not str — str on a categorical → numpy <U OOM
        g = gdf.groupby(cols, observed=True, dropna=False)["amount_mop"].agg(["sum", "size"]).reset_index()
        g = g.rename(columns={"sum": "amount_mop", "size": "筆數"})
        # 次數 = distinct sub-projects per (NG, V) — the KPI 數量/次數
        if {"ng_label", "vertical_label", "subproject"}.issubset(g.columns):
            g["次數"] = g.groupby(["ng_label", "vertical_label"])["subproject"].transform("nunique")
        return g.sort_values(["vertical_label", "amount_mop"] if "vertical_label" in g.columns
                             else ["amount_mop"], ascending=[True, False])

    # Write one Excel per year (12 files: 6 entity × {24, 25})
    for year in ("25", "24"):
        gc.collect()
        sub = _filter_year(df, ycol, year) if ycol else df.iloc[0:0]
        print(f"[{ent}] year {year}: {len(sub):,} rows", flush=True)
        if len(sub) == 0:
            print(f"[{ent}] year {year} empty — skip"); continue

        je_cols_present = [c for c in JE_KEEP if c in sub.columns]
        je_rows = sub[je_cols_present]
        h_break = _horizontal_breakdown(sub)
        v_break = _vertical_breakdown(sub)
        print(f"[{ent}] year {year} sheet rows — 大表: {len(je_rows):,} | "
              f"橫向: {len(h_break):,} | 縱向: {len(v_break):,}", flush=True)

        out_path = out_dir / f"{ent}_投資方向_{year}.xlsx"
        # NOTE: do NOT use xlsxwriter constant_memory here — at 624k rows on Windows it silently
        # dropped the 大表 rows (file 35.6MB → 3.3MB, sheet empty). Normal mode + the astype(object)
        # fix + categorical downcast keeps memory acceptable.
        with pd.ExcelWriter(out_path, engine="xlsxwriter") as w:
            idx = pd.DataFrame([
                ("0_index", "This map"),
                ("1_投資方向pivot", f"V × H cross-tab matrix for year {year} ({len(sub):,} rows, Σ amount_mop, NG header + 總計)"),
                ("2_橫向", f"Horizontal drill: (H, account_code, account_desc, project, subproject) × Σ amount ({len(h_break):,} aggregated rows)"),
                ("3_縱向", f"Vertical drill: (NG, V, project, subproject) × Σ amount ({len(v_break):,} aggregated rows)"),
                ("4_大表", f"ALL JE rows ({len(je_rows):,}) flat with classification + dimensions"),
            ], columns=["sheet", "contents"])
            idx.to_excel(w, sheet_name="0_index", index=False)

            kpi = _kpi_pivot(sub, categories)
            if not kpi.empty:
                # merge_cells=False: each row repeats ng_code/ng_label/V (no merged cells) so the
                # sheet is sortable directly in Excel. Rows already sorted NG0→NG11.
                kpi.to_excel(w, sheet_name="1_投資方向pivot", merge_cells=False)
            else:
                pd.DataFrame({"note": ["(no data)"]}).to_excel(w, sheet_name="1_投資方向pivot", index=False)

            h_break.to_excel(w, sheet_name="2_橫向", index=False)
            v_break.to_excel(w, sheet_name="3_縱向", index=False)

            if len(je_rows) <= XLSX_LIMIT:
                je_rows.to_excel(w, sheet_name="4_大表", index=False)
            else:
                n_chunks = (len(je_rows) + XLSX_LIMIT - 1) // XLSX_LIMIT
                for i in range(n_chunks):
                    chunk = je_rows.iloc[i*XLSX_LIMIT:(i+1)*XLSX_LIMIT]
                    chunk.to_excel(w, sheet_name=f"4_大表_p{i+1}of{n_chunks}", index=False)

        size_mb = out_path.stat().st_size / 1e6
        print(f"[{ent}] ✓ {out_path.name}  ({size_mb:.1f} MB, {time.time()-t0:.1f}s)", flush=True)
        written.append(out_path)
        # Free the per-year temporaries so the next year's subset copy doesn't fight fragmentation.
        del sub, je_rows, h_break, v_break
        gc.collect()
    return written


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--entity", choices=list(ENTITIES), default=None)
    args = p.parse_args()

    cat = yaml.safe_load(Path("conf/base/categories.yml").read_text(encoding="utf-8"))

    targets = [(args.entity, ENTITIES[args.entity])] if args.entity else list(ENTITIES.items())
    written = []
    for ent, com in targets:
        try:
            out = build(ent, com, cat)
            if out: written.append(out)
        except Exception as e:
            print(f"❌ {ent} failed: {e}", flush=True)

    print(f"\nDone — {len(written)} master audit xlsx written.", flush=True)
    for p in written:
        print(f"  {p}", flush=True)


if __name__ == "__main__":
    main()
