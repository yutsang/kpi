"""Prepare single Tableau-ready file combining all 6 entities' 25-year data.

Reads each entity's kpi_report.parquet, filters to year 25 (including 25_24SY
+ 25_23SY split-year buckets), adds entity column, concatenates, outputs:
  - tableau_combined_25.parquet (fast load in Tableau via Web Data Connector)
  - tableau_combined_25.csv (universal Tableau format)
  - tableau_combined_25.xlsx (multi-entity single sheet, may exceed 1M row limit)

Year tag column 'year_bucket' added: '25', '25_24SY', '25_23SY'
Entity column added: galaxy/sjm/wynn/vml/melco/mgm

Run:
  python scripts/prep_tableau_25.py
  python scripts/prep_tableau_25.py --format csv  # csv only (fastest)
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception: pass
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
import pandas as pd
import yaml

ENTITIES = {"galaxy":"company_1","sjm":"company_2","wynn":"company_3","vml":"company_4","melco":"company_5","mgm":"company_6"}

# 4 層 project 結構嘅細碼/細名來源欄 (diag_project_code 由原始 data book 確認)。
#   (subproject_code_col, subproject_name_col)；code=None 即原始檔冇比 dicj 更細嘅碼 → project_code 用 dicj-level。
# conf 同名 key (subproject_code_col / subproject_name_col) 可覆蓋呢度。
SUBPROJECT_COLS = {
    "galaxy": ("project_code", "subproject"),                          # B021 / Cross Border HK
    "vml":    ("Subproject", "SubProject_Name"),                       # SP00033 / Comprehensive Upgrade
    "melco":  ("Project & Sub-project ID", "project_mre"),             # 13c / SC Master Plan (乾淨細名)
    "mgm":    ("Project_code", "Project_name"),                        # 項目019-OPEX / 細名
    "sjm":    (None, "Subproject"),                                    # 冇細碼；細層只得名 (序號+名)
    "wynn":   (None, "Sub project"),                                   # 冇細碼；細層只得名
}

# Canonical V → NG0-NG11 mapping (Macau gaming framework)
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
    "V_PROPERTY_UPGRADE": ("NG11", "其他"),
    "V_OTHER": ("NG11", "其他"),
}


def _cn_kw(s) -> str:
    """中文 範疇 → NG keyword fallback (mirrors build_master_audit). 博彩 before 娛樂; 非博彩 = noise."""
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
    if not name:
        return None
    if name in df.columns:
        return name
    for c in df.columns:
        if str(c).strip() == str(name).strip():
            return c
    return None


import re as _re
_GCAND = [Path("HQ 投資方向_audit_0616.xlsx"), Path("data/HQ 投資方向_audit_0616.xlsx")]
_N2A = [(["galaxy", "銀河"], "galaxy"), (["wynn", "永利"], "wynn"), (["mgm", "美高梅"], "mgm"),
        (["melco", "新濠", "摩珀斯", "影匯", "影滙", "studio city", "city of dreams"], "melco"),
        (["sjm", "澳娛", "葡京", "回力", "上葡京"], "sjm"),
        (["威尼斯", "金沙", "sands", "londoner", "倫敦人", "parisian", "巴黎人", "venetian", "vml"], "vml")]


def _gname_alias(s):
    sl = str(s).strip().lower()
    for kws, a in _N2A:
        if any(k.lower() in sl for k in kws): return a
    return None


def _gndicj(v):
    v = str(v).strip()
    m = _re.match(r"^(博彩項目|項目)0*(\d.*)$", v)
    return m.group(1) + m.group(2) if m else v


def golden_name_map():
    """{alias: {dicj_norm: 項目名稱}} — 項目組要求 project name = golden DICJ 名稱 (per 承批公司)."""
    gp = next((p for p in _GCAND if p.exists()), None)
    if not gp:
        print("  ⚠ golden 檔揾唔到 → 唔 attach 項目名稱"); return {}
    try:
        g = pd.read_excel(gp, sheet_name="Database combine", dtype=str)
    except Exception as e:
        print(f"  ⚠ 讀 golden 失敗: {e}"); return {}
    g.columns = [str(c).strip() for c in g.columns]
    dcol = next((c for c in g.columns if c.strip() in ("DICJ Code", "DICJ")), None)
    ncol = next((c for c in g.columns if "項目名稱" in str(c) or "项目名称" in str(c)), None)
    acol = next((c for c in g.columns if c.strip() == "Amount"), None)
    acomp = next((c for c in g.columns if "承批" in str(c)), None)
    if not (dcol and ncol and acomp): return {}
    g["_a"] = g[acomp].map(_gname_alias)
    g["_d"] = g[dcol].astype(str).str.strip().map(_gndicj)
    g["_n"] = g[ncol].astype(str).str.strip()
    g["_amt"] = pd.to_numeric(g[acol].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0.0) if acol else 0.0
    gg = g[g["_n"].ne("") & g["_n"].ne("nan") & g["_a"].notna()]
    out = {}
    for a in gg["_a"].unique():
        sub = gg[gg["_a"] == a]
        idx = sub.groupby("_d")["_amt"].apply(lambda s: s.abs().idxmax())   # 每 dicj 攞金額最大嗰個名
        out[a] = dict(zip(sub.loc[idx, "_d"], sub.loc[idx, "_n"]))
    return out


def run(fmt="csv", out_dir="data/tableau"):
    """Build Tableau files. DEFAULT = csv (one combined tableau_combined_25.csv). Other formats kept
    for the kedro generate/tableau pipelines. Importable so they can call it (no argparse)."""
    from kpi.lib.conf import load_categories
    from kpi.pipelines.step2_tag_projects._logic import normalize_ng_code
    cats = load_categories()
    gname = golden_name_map()
    if gname:
        print(f"  golden 名 map: {', '.join(f'{a}={len(m)}' for a, m in gname.items())}")
    frames = []
    for ent, com in ENTITIES.items():
        parquet = Path(f"data/{ent}/output/{com}_kpi_report.parquet")
        if not parquet.exists():
            print(f"⚠️  {ent}: {parquet} missing — skip")
            continue

        cfg = yaml.safe_load(Path(f"conf/{com}/parameters.yml").read_text(encoding="utf-8"))
        amt_cfg = cfg.get("columns",{}).get("amount","")
        df = pd.read_parquet(parquet)
        # Fuzzy amt col
        if amt_cfg not in df.columns:
            for c in df.columns:
                if c.strip()==amt_cfg.strip(): amt_cfg=c; break

        # Filter to year 23 + 24 + 25 buckets (all three delivery years for Tableau)
        ycol = next((c for c in ("report_period","report_year","Yr related","years") if c in df.columns), None)
        if ycol:
            s = df[ycol].astype(str)
            df = df[s.str.startswith("23") | s.str.startswith("24") | s.str.startswith("25")
                    | (s == "Yr 2023") | (s == "Yr 2024") | (s == "Yr 2025")].copy()
        if len(df)==0:
            print(f"⚠️  {ent}: 0 rows for year 23/24/25 — skip"); continue

        # Add columns
        df["entity"] = ent
        # Year bucket from report_period: '25' / '25_24SY' / '25_23SY' / '24' / '24_23SY' etc.
        if ycol == "report_period":
            df["year_bucket"] = df[ycol].astype(str)
        elif ycol:
            df["year_bucket"] = df[ycol].astype(str).str[:2]  # take first 2 chars as year
        else:
            df["year_bucket"] = "?"

        # Normalize amount column name — always create amount_mop
        if amt_cfg and amt_cfg in df.columns and amt_cfg != "amount_mop":
            df["amount_mop"] = pd.to_numeric(df[amt_cfg], errors="coerce").fillna(0)
        elif "amount_mop" not in df.columns:
            # Fallback: try common amount col names
            for cand in ["amount_mop", "amount", "Amount", "MOP Amt", "Reported Amount(MOP)",
                          "Entry Voucher Amount/ Expense Amount", "Entry Voucher Amount/ Expense Amount ",
                          "amount_mop_split", "Amount - Amended"]:
                if cand in df.columns:
                    df["amount_mop"] = pd.to_numeric(df[cand], errors="coerce").fillna(0)
                    print(f"  [{ent}] fallback amount col: '{cand}'")
                    break
            else:
                print(f"  ⚠️  [{ent}] no amount col found! cols={list(df.columns)[:10]}")
                df["amount_mop"] = 0.0

        # Keep ONLY essential cols (compact for Tableau)
        keep = ["entity", "year_bucket"]
        for c in ["amount_mop", "horizontal_id", "horizontal_label", "vertical_id",
                  "vertical_label", "ng_scope", "final_capex_opex", "row_type"]:
            if c in df.columns: keep.append(c)
        # NG — ALWAYS the project team's databook 範疇 column (項目性質/項目類型/ng_theme/Section.1/NG11
        # Category), NEVER derived from V. Mirrors build_master_audit so 大表 + Tableau NG match exactly.
        # (Only V + H are OUR classification work; NG must stay the project team's.) Unmapped → (未分類).
        _ng_cols = []
        for _nm in ([cfg.get("columns", {}).get("ng11_category", "")]
                    + [(ys.get("columns_override") or {}).get("ng11_category")
                       for ys in (cfg.get("yearly_sources") or [])]):
            _fc = _fuzzy_col(df, _nm)
            if _fc and _fc not in _ng_cols:
                _ng_cols.append(_fc)
        for _c in df.columns:
            if _c not in _ng_cols and any(k in str(_c) for k in
                    ("項目性質", "項目類型", "項目分類", "範疇", "NG11 Category", "NG Category")):
                _ng_cols.append(_c)

        def _resolve_ng(x):
            for cand in (x, str(x).upper().replace(" ", "")):
                r = normalize_ng_code(cand, cats) or ""
                if r[:2] == "NG" and r[2:].isdigit():
                    return r
            return _cn_kw(x)
        _ngc = pd.Series("", index=df.index, dtype="object")
        for _fc in _ng_cols:
            _m = {x: _resolve_ng(x) for x in set(df[_fc].astype(str).unique())}
            _r = df[_fc].astype(str).map(_m).fillna("")
            _r = _r.where(_r.str.fullmatch(r"NG\d+").fillna(False), "")
            _ngc = _ngc.mask(_ngc.eq(""), _r)
        df["ng_code"] = _ngc
        _nglab = {ng: lbl for ng, lbl in V_TO_NG.values()}
        df["ng_label"] = df["ng_code"].map(lambda n: _nglab.get(str(n), "(未分類)"))
        keep.append("ng_code"); keep.append("ng_label")
        _mp = int(_ngc.str.fullmatch(r"NG\d+").fillna(False).sum())
        print(f"  [{ent}] NG from databook {_ng_cols}: {_mp:,}/{len(df):,} mapped (NG=項目組, 非V)")
        # Per-entity native cols (preserve project + subproject + acct for drill-down)
        proj_col = cfg.get("columns",{}).get("project","")
        ac_col = cfg.get("columns",{}).get("account_code","")
        ad_col = cfg.get("columns",{}).get("account_desc","")
        dn_col = cfg.get("columns",{}).get("description","")
        vd_col = cfg.get("columns",{}).get("vendor","")
        # subproject 細名欄：conf > SUBPROJECT_COLS dict > 候選
        _dict_code, _dict_name = SUBPROJECT_COLS.get(ent, (None, None))
        _want_name = cfg.get("subproject_name_col") or _dict_name
        sub_col = (_want_name if _want_name in df.columns
                   else next((c for c in ("Sub project", "SubProject_Name", "Subproject_Name",
                                          "subproject", "项目名称中文", "项目英文名称",
                                          "Initiative Name", "Contents Name") if c in df.columns), None))
        for src, tgt in [(proj_col,"project"),(sub_col,"subproject"),
                          (ac_col,"account_code"),(ad_col,"account_desc"),
                          (dn_col,"description"),(vd_col,"vendor")]:
            if src and src in df.columns:
                df[tgt] = df[src].astype(str).str.strip().str.replace(r"[\r\n]+", " ", regex=True)
                keep.append(tgt)

        # ── UNIFORM extra detail layers (audit_detail_cols) — same set as the audit 大表 so Tableau can
        #    drill每條數 at consistent granularity (科目層級/科目明細/發票號/PO號/成本中心/WBS子項/憑證號).
        _adc = cfg.get("audit_detail_cols") or {}
        _default_raw = {"項目組H": "pt_class_H", "項目組V": "pt_class_V"}   # unified-raw reference labels
        for _name in ("科目層級", "科目明細", "發票號", "PO號", "成本中心", "WBS子項", "憑證號", "項目組H", "項目組V"):
            _raws = _adc.get(_name, "") or _default_raw.get(_name, "")
            _raws = _raws if isinstance(_raws, list) else ([_raws] if _raws else [])
            _ser = pd.Series("", index=df.index, dtype=object)
            for _r in _raws:
                _src = (_r if _r in df.columns
                        else next((c for c in df.columns if str(c).strip() == str(_r).strip()), None))
                if _src:
                    _v = df[_src].astype(str).str.strip().str.replace(r"[\r\n]+", " ", regex=True)
                    _ser = _ser.where(~_ser.isin(["", "nan", "None"]), _v)
            df[_name] = _ser.replace({"nan": "", "None": ""})
            keep.append(_name)

        # unified-raw extra cols (user 2026-06-14: carry every column incl remarks into Tableau)
        for _u in ("project_code", "dicj_code", "adjustment_amount", "adjusted_amount",
                   "adjust_lv1", "adjust_lv2", "source", "comp_type", "is_labor", "is_internal",
                   "take_flag", "take_flag2", "netoff_flag", "internal", "remark",
                   "調整金額", "調整後金額", "調整一級", "調整二級"):
            if _u in df.columns and _u not in keep:
                keep.append(_u)

        # Add merged project_full = "project | subproject | description" for Tableau display
        merge_parts = []
        for c in ("project", "subproject", "description"):
            if c in df.columns:
                merge_parts.append(df[c].astype(str).fillna(""))
        if merge_parts:
            df["project_full"] = merge_parts[0]
            for part in merge_parts[1:]:
                df["project_full"] = df["project_full"] + " | " + part
            keep.append("project_full")

        # ── project_code = 純細碼 (subproject code) ──
        # per-entity conf `subproject_code_col` 指明真細碼欄：galaxy project_code(B021) / vml Subproject(SP00033) /
        #   melco 'Project & Sub-project ID'(13c) / mgm Project_code(項目019-OPEX)。空位 fill dicj_code。
        # sjm/wynn 原始檔冇比 dicj 更細嘅碼 → 唔設 subproject_code_col → project_code = dicj_code(細層只得名，喺 subproject 欄)。
        _subcode = cfg.get("subproject_code_col") or _dict_code
        _dicj = (df["dicj_code"].astype("string").fillna("").str.strip()
                 if "dicj_code" in df.columns else pd.Series("", index=df.index))
        if _subcode and _subcode in df.columns:
            _pc = df[_subcode].astype("string").fillna("").str.strip()
            _pc = _pc.mask(_pc.eq("") | _pc.isin(["nan", "None"]), _dicj)
            _srcdesc = f"conf:{_subcode}+dicj"
        else:
            _has_native = "project_code" in df.columns
            _pc = ((df["project_code"] if _has_native else pd.Series("", index=df.index))
                   .astype("string").fillna("").str.strip())
            for _fb in (["dicj_code"] if _has_native else ["project", "dicj_code"]):
                if _fb in df.columns:
                    _pc = _pc.mask(_pc.eq("") | _pc.isin(["nan", "None"]),
                                   df[_fb].astype("string").fillna("").str.strip())
            _srcdesc = ("native+dicj" if _has_native else "project碼+dicj")
            if _subcode:
                _srcdesc += f" (⚠ conf subproject_code_col='{_subcode}' 唔喺 df)"
        df["project_code"] = _pc.replace({"nan": "", "None": ""})
        if "project_code" not in keep:
            keep.append("project_code")
        _nb = int(df["project_code"].astype(str).str.strip().isin(["", "nan", "None"]).sum())
        print(f"  [{ent}] project_code ← {_srcdesc}, blank={_nb}")

        # ── 項目名稱 = golden DICJ 名稱 (by dicj_code) — 項目組要求 project name = golden 名（roll-up 層）──
        # umbrella (golden 冇單一名，如 sjm 項目40=體育盛事 多個事件) → blank 用我哋 subproject 名頂上。
        _gm = gname.get(ent, {})
        if _gm and "dicj_code" in df.columns:
            df["項目名稱"] = df["dicj_code"].astype(str).fillna("").str.strip().map(_gndicj).map(_gm).fillna("")
            if "subproject" in df.columns:
                _bl = df["項目名稱"].astype(str).str.strip().isin(["", "nan", "None"])
                df.loc[_bl, "項目名稱"] = df.loc[_bl, "subproject"].astype(str).str.strip()
            keep.append("項目名稱")
            _gn_blank = int(df["項目名稱"].astype(str).str.strip().isin(["", "nan", "None"]).sum())
            print(f"  [{ent}] 項目名稱 ← golden(+subproject fallback): 仍 blank {_gn_blank:,}/{len(df):,}")

        # ── H/V label 跟最新 categories.yml (id→label) re-map：改名(藝人演出費/Comp房間…)唔使重跑 pipeline ──
        _hl = {h["id"]: h["label"] for h in (cats.get("horizontals") or [])}
        _vl = {v["id"]: v["label"] for v in (cats.get("verticals") or [])}
        if "horizontal_id" in df.columns and "horizontal_label" in df.columns and _hl:
            _nl = df["horizontal_id"].astype(str).map(_hl)
            df["horizontal_label"] = _nl.where(_nl.notna(), df["horizontal_label"])
        if "vertical_id" in df.columns and "vertical_label" in df.columns and _vl:
            _nv = df["vertical_id"].astype(str).map(_vl)
            df["vertical_label"] = _nv.where(_nv.notna(), df["vertical_label"])

        # ── 調整前 / 調整 / 調整後 (萬元) — 項目組對數用。調整後 = native 調整後金額 else amount；調整前 = 調整後 − 調整金額 ──
        _adj = (pd.to_numeric(df["調整金額"], errors="coerce").fillna(0.0)
                if "調整金額" in df.columns else pd.Series(0.0, index=df.index))
        _post = (pd.to_numeric(df["調整後金額"], errors="coerce")
                 if "調整後金額" in df.columns else pd.Series(pd.NA, index=df.index, dtype="Float64"))
        _post = pd.to_numeric(_post, errors="coerce").fillna(df["amount_mop"])
        df["調整後_萬"] = _post / 1e4
        df["調整_萬"] = _adj / 1e4
        df["調整前_萬"] = (_post - _adj) / 1e4
        keep += ["調整前_萬", "調整_萬", "調整後_萬"]

        sub = df[keep].copy()
        # Ensure amount_mop col exists in sub
        if "amount_mop" not in sub.columns:
            sub["amount_mop"] = 0.0
        frames.append(sub)
        print(f"✓ {ent}: {len(sub):,} rows, {sub['amount_mop'].sum()/1e6:.0f}M MOP")

    if not frames:
        print("❌ no data"); return

    combined = pd.concat(frames, ignore_index=True)
    # year_bucket keeps the FULL 5 split-year buckets (24 / 24_23SY / 25 / 25_23SY / 25_24SY) as text,
    # matching the data\review 大表's per-bucket pivots. (No numeric 'year' col — it confuses Tableau typing.)
    combined["year_bucket"] = combined["year_bucket"].astype(str)
    print(f"\nCombined: {len(combined):,} rows, {combined['amount_mop'].sum()/1e6:.0f}M total")
    print(f"year_bucket values: {sorted(combined['year_bucket'].unique())}")

    # ── 4 層 project 欄改清晰名 (方便查睇) ──
    #   dicj code        = DICJ 碼 (粗)
    #   project          = DICJ 大項目名 (= golden 項目名稱)；原生 project 名欄 drop 走由佢接手
    #   subproject code  = 細碼 (galaxy B021 / vml SP00033 / melco 13c / mgm 項目019-OPEX；sjm/wynn=dicj層)
    #   subproject       = 細名
    if "項目名稱" in combined.columns:
        combined = combined.drop(columns=[c for c in ["project"] if c in combined.columns])
        combined = combined.rename(columns={"項目名稱": "project"})
    combined = combined.rename(columns={"dicj_code": "dicj code", "project_code": "subproject code"})
    print(f"Cols: {list(combined.columns)}")

    # ── cube / cube-detail: ONE aggregated file = cross-tab source (no union, no stitching) ──
    if fmt in ("cube", "cube-detail"):
        dims = ["entity", "year_bucket", "ng_code", "ng_label",
                "vertical_id", "vertical_label", "horizontal_id", "horizontal_label",
                "ng_scope", "final_capex_opex"]
        if fmt == "cube-detail":   # keep drill-down dims (4 層 project + account / vendor)
            dims += ["dicj code", "project", "subproject code", "subproject", "account_code", "account_desc", "vendor"]
        dims = [c for c in dims if c in combined.columns]
        cube = (combined.groupby(dims, dropna=False, observed=True)["amount_mop"]
                        .agg(amount_mop="sum", n_rows="size").reset_index())
        _od = Path(out_dir); _od.mkdir(parents=True, exist_ok=True)
        stem = "tableau_cube_detail" if fmt == "cube-detail" else "tableau_cube"
        p = _od / f"{stem}.csv"
        cube.to_csv(p, index=False, encoding="utf-8-sig")
        print(f"✓ wrote {p}  (Tableau → Text File)  — {len(cube):,} rows, {cube['amount_mop'].sum()/1e6:.0f}M")
        if len(cube) <= 1_048_574:
            px = _od / f"{stem}.xlsx"
            cube.to_excel(px, index=False, engine="xlsxwriter")
            print(f"✓ wrote {px}  (Tableau → Microsoft Excel)")
        else:
            print(f"  ⚠ {len(cube):,} rows > Excel 1M limit → use the .csv via Text File connector.")
        return

    # Default: 6 per-entity × 2 year Excels (12 files for Tableau union)
    # Write to data/{ent}/output/tableau_{yr}_{ent}.xlsx so they sit alongside
    # the entity's kpi_report.parquet.
    if fmt == "per-entity-xlsx":
        XLSX_LIM = 1_048_574
        _od = Path(out_dir); _od.mkdir(parents=True, exist_ok=True)
        for (ent, yr), sub in combined.groupby(["entity", combined["year_bucket"].astype(str).str[:2]]):
            p = _od / f"tableau_{yr}_{ent}.xlsx"
            if len(sub) <= XLSX_LIM:
                sub.to_excel(p, index=False, engine="xlsxwriter")
                print(f"✓ wrote {p} ({p.stat().st_size/1e6:.1f} MB, {len(sub):,} rows)")
            else:
                n_chunks = (len(sub) + XLSX_LIM - 1) // XLSX_LIM
                for i in range(n_chunks):
                    chunk = sub.iloc[i*XLSX_LIM:(i+1)*XLSX_LIM]
                    pc = _od / f"tableau_{yr}_{ent}_p{i+1}of{n_chunks}.xlsx"
                    chunk.to_excel(pc, index=False, engine="xlsxwriter")
                    print(f"✓ wrote {pc} ({pc.stat().st_size/1e6:.1f} MB, {len(chunk):,} rows)")
        return

    # ── csv-per-entity: 6 row-level CSVs (every JE line + all dims incl description) ──
    if fmt == "csv-per-entity":
        _od = Path(out_dir); _od.mkdir(parents=True, exist_ok=True)
        for ent, sub in combined.groupby("entity"):
            p = _od / f"tableau_detail_{ent}.csv"
            sub.to_csv(p, index=False, encoding="utf-8-sig")
            print(f"✓ {p}  — {len(sub):,} rows, {sub['amount_mop'].sum()/1e6:.0f}M (Tableau → Text File)")
        return

    out_base = Path("tableau_combined_25")
    if fmt in ("all","parquet"):
        combined.to_parquet(out_base.with_suffix(".parquet"), index=False)
        print(f"✓ wrote {out_base}.parquet ({out_base.with_suffix('.parquet').stat().st_size/1e6:.1f} MB)")
    if fmt in ("all","csv"):
        combined.to_csv(out_base.with_suffix(".csv"), index=False, encoding="utf-8-sig")
        print(f"✓ wrote {out_base}.csv ({out_base.with_suffix('.csv').stat().st_size/1e6:.1f} MB)")
    if fmt in ("all","xlsx"):
        XLSX_LIM = 1_048_574
        if len(combined) <= XLSX_LIM:
            combined.to_excel(out_base.with_suffix(".xlsx"), index=False, engine="xlsxwriter")
            print(f"✓ wrote {out_base}.xlsx ({out_base.with_suffix('.xlsx').stat().st_size/1e6:.1f} MB)")
        else:
            for ent, sub in combined.groupby("entity"):
                p = Path(f"tableau_combined_25_{ent}.xlsx")
                sub.to_excel(p, index=False, engine="xlsxwriter")
                print(f"✓ wrote {p.name} ({p.stat().st_size/1e6:.1f} MB, {len(sub):,} rows)")

def main():
    p = argparse.ArgumentParser()
    p.add_argument("--format", choices=["cube", "cube-detail", "csv", "csv-per-entity", "all", "parquet", "xlsx", "per-entity-xlsx"],
                   default="csv",
                   help="DEFAULT csv = ONE row-level tableau_combined_25.csv (all dims incl 調整/remark). "
                        "Others kept but rarely needed: cube=aggregated; csv-per-entity=6 files; per-entity-xlsx=old union")
    p.add_argument("--out", default="data/tableau", help="output dir for per-entity-xlsx")
    args = p.parse_args()
    run(args.format, args.out)


if __name__ == "__main__":
    main()
