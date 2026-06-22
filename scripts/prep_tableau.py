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
    "melco":  ("Project & Sub-project ID", "Project/Sub-Project Name_1"),  # 13c / Water park operations (100% 乾淨細名)
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
    """Return (gname, gname_ng):
       gname    = {alias: {dicj_norm: 項目名稱}}            ← 每 dicj idxmax（fallback）
       gname_ng = {alias: {(dicj_norm, NGcode): 項目名稱}}  ← per (DICJ,NG) 正解
    同一 DICJ 嘅博彩(NG0)/非博彩(NG1-11)係兩個唔同項目名；舊 idxmax 只攞金額大嗰個 → 非博彩行冚博彩名。
    用 golden「投資範疇」欄 → _cn_kw → NG code，回填時每行用 (dicj, ng_code) 對返正確名。"""
    gp = next((p for p in _GCAND if p.exists()), None)
    if not gp:
        print("  ⚠ golden 檔揾唔到 → 唔 attach 項目名稱"); return {}, {}
    try:
        g = pd.read_excel(gp, sheet_name="Database combine", dtype=str)
    except Exception as e:
        print(f"  ⚠ 讀 golden 失敗: {e}"); return {}, {}
    g.columns = [str(c).strip() for c in g.columns]
    dcol = next((c for c in g.columns if c.strip() in ("DICJ Code", "DICJ")), None)
    ncol = next((c for c in g.columns if "項目名稱" in str(c) or "项目名称" in str(c)), None)
    acol = next((c for c in g.columns if c.strip() == "Amount"), None)
    acomp = next((c for c in g.columns if "承批" in str(c)), None)
    ngcol = next((c for c in g.columns if c.strip() == "投資範疇"), None)  # NG theme（≠「投資範疇|是否博彩」）
    if not (dcol and ncol and acomp): return {}, {}
    g["_a"] = g[acomp].map(_gname_alias)
    g["_d"] = g[dcol].astype(str).str.strip().map(_gndicj)
    g["_n"] = g[ncol].astype(str).str.strip()
    g["_ng"] = g[ngcol].astype(str).map(_cn_kw) if ngcol else ""
    g["_amt"] = pd.to_numeric(g[acol].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0.0) if acol else 0.0
    gg = g[g["_n"].ne("") & g["_n"].ne("nan") & g["_a"].notna()]
    out, out_ng = {}, {}
    for a in gg["_a"].unique():
        sub = gg[gg["_a"] == a]
        idx = sub.groupby("_d")["_amt"].apply(lambda s: s.abs().idxmax())   # 每 dicj 攞金額最大嗰個名
        out[a] = dict(zip(sub.loc[idx, "_d"], sub.loc[idx, "_n"]))
        if ngcol is not None:
            sub2 = sub[sub["_ng"].ne("")]
            if len(sub2):
                idx2 = sub2.groupby(["_d", "_ng"])["_amt"].apply(lambda s: s.abs().idxmax())
                out_ng[a] = {k: sub2.loc[i, "_n"] for k, i in idx2.items()}
    return out, out_ng


def run(fmt="csv", out_dir="data/tableau"):
    """Build Tableau files. DEFAULT = csv (one combined tableau_combined_25.csv). Other formats kept
    for the kedro generate/tableau pipelines. Importable so they can call it (no argparse)."""
    from kpi.lib.conf import load_categories
    from kpi.pipelines.step2_tag_projects._logic import normalize_ng_code
    cats = load_categories()
    gname, gname_ng = golden_name_map()
    if gname:
        print(f"  golden 名 map: {', '.join(f'{a}={len(m)}' for a, m in gname.items())}"
              f"  | (dicj,NG) keys: {', '.join(f'{a}={len(m)}' for a, m in gname_ng.items())}")
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
        # wynn：原生細欄「subproject」(真細名 Music Concerts… 6%) 比 sub_col「Sub project」(=項目名 92%)
        #   更細，但 sub_col 揀咗 Sub project 會覆蓋佢 → 先捕捉，for-loop 後 overlay 返（有就用更細）。
        _fine_sub = (df["subproject"].astype(str).str.strip().str.replace(r"[\r\n]+", " ", regex=True)
                     if (ent == "wynn" and "subproject" in df.columns and sub_col != "subproject") else None)
        for src, tgt in [(proj_col,"project"),(sub_col,"subproject"),
                          (ac_col,"account_code"),(ad_col,"account_desc"),
                          (dn_col,"description"),(vd_col,"vendor")]:
            if src and src in df.columns:
                df[tgt] = df[src].astype(str).str.strip().str.replace(r"[\r\n]+", " ", regex=True)
                keep.append(tgt)
        if _fine_sub is not None and "subproject" in df.columns:
            _nb = ~_fine_sub.isin(["", "nan", "None", "NaN", "<NA>"])
            if int(_nb.sum()):
                df.loc[_nb, "subproject"] = _fine_sub[_nb]
                print(f"  [wynn subproject] overlay 更細 native subproject {int(_nb.sum()):,} 行（其餘用 Sub project）")

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
        _gmng = gname_ng.get(ent, {})
        if (_gm or _gmng) and "dicj_code" in df.columns:
            _dn = df["dicj_code"].astype(str).fillna("").str.strip().map(_gndicj)
            _ngc = df["ng_code"].astype(str).str.strip() if "ng_code" in df.columns else pd.Series("", index=df.index)
            # 每行先用 (dicj, ng_code) 對返正確名（修博彩/非博彩冚錯）；冇就 dicj idxmax fallback
            _nm = [(_gmng.get((d, n)) or _gm.get(d, "")) for d, n in zip(_dn.tolist(), _ngc.tolist())]
            df["項目名稱"] = pd.Series(_nm, index=df.index).fillna("")
            _ng_used = sum(1 for d, n in zip(_dn.tolist(), _ngc.tolist()) if (d, n) in _gmng)
            if "subproject" in df.columns:
                _bl = df["項目名稱"].astype(str).str.strip().isin(["", "nan", "None"])
                df.loc[_bl, "項目名稱"] = df.loc[_bl, "subproject"].astype(str).str.strip()
            keep.append("項目名稱")
            _gn_blank = int(df["項目名稱"].astype(str).str.strip().isin(["", "nan", "None"]).sum())
            print(f"  [{ent}] 項目名稱 ← golden(dicj,NG對 {_ng_used:,}行)(+subproject fallback): 仍 blank {_gn_blank:,}/{len(df):,}")

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

    # ── H_UTILITY → H_MAINTENANCE（user: 水電/utility 放維護費，唔好獨立 utility 類）──
    if "horizontal_id" in combined.columns:
        _um = combined["horizontal_id"].astype(str).str.strip().eq("H_UTILITY")
        if int(_um.sum()):
            combined.loc[_um, "horizontal_id"] = "H_MAINTENANCE"
            if "horizontal_label" in combined.columns:
                combined.loc[_um, "horizontal_label"] = "維護費"
            print(f"  [H_UTILITY→維護費] {int(_um.sum()):,} 行")

    # ── capex 只能入 {建設 / 設施器具 / 人工}（user 2026-06-18）──
    #   (a) 明確 account：CIP/裝修/Renovation/新工作范圍 → 建設（即使現標人工）
    #   (b) 通用：capex H ∉ {建設,器具,人工} → 採購/供應/設備類=器具；其餘(show/WIP/其他)=建設(舞台亦建設)
    if "horizontal_id" in combined.columns and "final_capex_opex" in combined.columns and "account_desc" in combined.columns:
        _cap = combined["final_capex_opex"].astype(str).str.strip().eq("Capex")
        _acc = combined["account_desc"].astype(str)
        _hid = combined["horizontal_id"].astype(str).str.strip()
        _CON = ["CIP-A&A", "CIP-OTHER", "CIP-", "Deposits paid - Renovation", "Renovation (WBS)",
                "租賃物業裝修", "物業裝修", "新工作范圍", "新工作範圍"]
        _EQKW = ["PURCHASES", "採購", "Par Stock", "Stock", "Inventory", "Supplies", "China", "Glass",
                 "Chinaware", "FF&E", "設備", "器具", "Equipment", "Software", "License", "Kitchen",
                 "FA -", "O/E-", "廠房和設備", "Plant and Equipment", "Food", "Beverage"]
        _ALLOW = {"H_CONSTRUCTION", "H_EQUIP", "H_LABOR"}
        _eqkw = _acc.apply(lambda s: any(k in str(s) for k in _EQKW))
        _con0 = _cap & _acc.apply(lambda s: any(k in str(s) for k in _CON))   # 明確建設(即使現人工)
        _bad = _cap & ~_hid.isin(_ALLOW)                                      # H 唔喺 allow
        _to_eqp = _cap & _eqkw & ~_con0          # capex 採購/設備類 → 器具（即使現標人工，如 PURCHASES）
        _to_con = _con0 | (_bad & ~_eqkw)        # CIP/裝修 + 其餘非 allow → 建設
        for _m, _h, _lab, _nm in [(_to_eqp, "H_EQUIP", "設施及器具採購", "器具"),
                                   (_to_con, "H_CONSTRUCTION", "建設與設施支出", "建設")]:
            if int(_m.sum()):
                combined.loc[_m, "horizontal_id"] = _h
                if "horizontal_label" in combined.columns:
                    combined.loc[_m, "horizontal_label"] = _lab
                print(f"  [capex→{_nm}] {int(_m.sum()):,} 行")

    # ── OTA / 旅行社 / 佣金 → 廣告及推廣（user 2026-06-18：呢類係推廣開支，唔係其他）──
    if "horizontal_id" in combined.columns and "account_desc" in combined.columns:
        def _is_ota(s):
            s = str(s)
            if any(k in s for k in ("佣金", "旅行社", "OTA")):
                return True
            return ("Commission" in s) and ("Agent" in s or "Travel" in s)
        _otam = combined["account_desc"].astype(str).apply(_is_ota) & \
            combined["horizontal_id"].astype(str).str.strip().isin(["H_OTHER", "H_PROFESSIONAL"])
        if int(_otam.sum()):
            combined.loc[_otam, "horizontal_id"] = "H_ADVERTISING"
            if "horizontal_label" in combined.columns:
                combined.loc[_otam, "horizontal_label"] = "廣告及推廣"
            print(f"  [OTA/佣金→廣告] {int(_otam.sum()):,} 行")

    # ── vml H 矛盾修正（compare_h_ref: 項目組H 確認我哋錯）：OUTSIDE SERV→專業(opex) / SALARIES→人工 / CONTRACT LABOR→人工 ──
    if "entity" in combined.columns and "account_desc" in combined.columns and "horizontal_id" in combined.columns:
        _vml = combined["entity"].astype(str).eq("vml")
        _acc2 = combined["account_desc"].astype(str)
        _h2 = combined["horizontal_id"].astype(str).str.strip()
        _opex = ~combined["final_capex_opex"].astype(str).str.strip().eq("Capex") if "final_capex_opex" in combined.columns else True
        _m1 = _vml & _opex & _acc2.str.contains("OUTSIDE SERV", case=False, na=False) & _h2.eq("H_OTHER")
        _m2 = _vml & _acc2.str.contains("SALARIES|Salaries|Salary", case=False, regex=True, na=False) & _h2.eq("H_OTHER")
        _m3 = _vml & _acc2.str.contains("CONTRACT LABOR|Contract Labor", case=False, regex=True, na=False) & _h2.isin(["H_PROFESSIONAL", "H_OTHER"])
        _m4 = _vml & _opex & _acc2.str.contains("SOFTWARE & HOSTING|SOFTWARE AND HOSTING", case=False, regex=True, na=False) & _h2.eq("H_OTHER")
        _m5 = _vml & _acc2.str.contains("SOCIAL SECURITY|Social Security", case=False, regex=True, na=False) & _h2.eq("H_OTHER")
        # _m3 CONTRACT LABOR → 專業（user 2026-06-19：vml 人工 over 主因，合約勞工係專業服務唔係內部 staff）
        for _m, _hid, _lab in [(_m1, "H_PROFESSIONAL", "專業服務費"), (_m2, "H_LABOR", "人工成本"), (_m3, "H_PROFESSIONAL", "專業服務費"),
                               (_m4, "H_PROFESSIONAL", "專業服務費"), (_m5, "H_LABOR", "人工成本")]:
            if int(_m.sum()):
                combined.loc[_m, "horizontal_id"] = _hid
                combined.loc[_m, "horizontal_label"] = _lab
                print(f"  [vml H修正→{_lab}] {int(_m.sum()):,} 行")

        # ── vml 從 人工 剷走 cost-allocation/contract/outside-serv/travel（HQ staff 細；user 2026-06-19 acct dump）──
        _vlab = _vml & combined["horizontal_id"].astype(str).str.strip().eq("H_LABOR")
        _vcode = combined["account_code"].astype(str).str.replace(r"\.0$", "", regex=True) if "account_code" in combined.columns else pd.Series("", index=combined.index)
        _va = combined["account_desc"].astype(str)
        _v_alloc = _vlab & _vcode.str.startswith("89000")                              # 89000 部門間分攤 → 其他
        _v_con = _vlab & _va.str.contains("CONTRACT LABOR|Contract Labor", case=False, regex=True, na=False)  # → 專業
        _v_os = _vlab & _va.str.contains("OUTSIDE SERV", case=False, na=False)         # → 專業
        _v_tr = _vlab & _va.str.contains("TRAVEL & ENT|TRAVEL&ENT", case=False, regex=True, na=False)  # → 其他
        for _m, _hid, _lab in [(_v_alloc, "H_OTHER", "其他"), (_v_con, "H_PROFESSIONAL", "專業服務費"),
                               (_v_os, "H_PROFESSIONAL", "專業服務費"), (_v_tr, "H_OTHER", "其他")]:
            if int(_m.sum()):
                combined.loc[_m, "horizontal_id"] = _hid
                combined.loc[_m, "horizontal_label"] = _lab
        print(f"  [vml 人工剷] 89000分攤→其他={int(_v_alloc.sum())} | Contract→專業={int(_v_con.sum())} | OutSrv→專業={int(_v_os.sum())} | Travel→其他={int(_v_tr.sum())}")

    # ── galaxy 部門分攤修正（user 2026-06-18 逐項定）：EVS(清潔)→維護費；Allocation-Entertainment→廣告 ──
    #   Commission 維持其他（信用卡手續費,非中介佣金）；Allocation-R&M 由通用救援 R&M token 接手→維護。
    if "entity" in combined.columns and "account_desc" in combined.columns and "horizontal_id" in combined.columns:
        _gal = combined["entity"].astype(str).eq("galaxy")
        _accg = combined["account_desc"].astype(str)
        _hg = combined["horizontal_id"].astype(str).str.strip().eq("H_OTHER")
        _g_evs = _gal & _hg & _accg.str.contains("Allocation", case=False, na=False) & _accg.str.contains("EVS", na=False)
        _g_ent = _gal & _hg & _accg.str.contains("Allocation", case=False, na=False) & _accg.str.contains("Entertainment", case=False, na=False)
        for _m, _hid, _lab, _nm in [(_g_evs, "H_MAINTENANCE", "維護費", "EVS→維護"),
                                     (_g_ent, "H_ADVERTISING", "廣告及推廣", "Entertainment分攤→廣告")]:
            if int(_m.sum()):
                combined.loc[_m, "horizontal_id"] = _hid
                if "horizontal_label" in combined.columns:
                    combined.loc[_m, "horizontal_label"] = _lab
                print(f"  [galaxy {_nm}] {int(_m.sum()):,} 行")

    # ── 其他(H_OTHER) 救援：account_desc 明確屬某真 H bucket 嘅搬走（user 2026-06-18 逐家睇 其他 dump）──
    #   tie-safe（H 唔喺 golden tie key）；只郁 H_OTHER；first-match-wins；收入/回收/contra/分攤 留 其他。
    #   留喺其他唔郁：匯兌/FX/利息/稅/Cost Recovery/Interco 分攤/Adjustment/Variance/Prepaid（無 token 自然唔搬）。
    if "horizontal_id" in combined.columns and "account_desc" in combined.columns:
        _accU = combined["account_desc"].astype(str).str.upper()
        _hO = combined["horizontal_id"].astype(str).str.strip().eq("H_OTHER")
        # guard：呢類就算撞 token 都唔搬（收入/回收/contra 抵銷項）
        _keepO = _accU.str.contains("REVENUE|CONTRA|RECOVERY", regex=True, na=False)
        # (target_id, label, [account_desc UPPER token]) — 順序 = 優先（前者先 match）
        _RESCUE = [
            ("H_HOTEL_ROOM", "Comp房間", ["COMPLIMENTARY ROOM"]),
            ("H_FNB", "Comp餐飲", ["COMPLIMENTARY FOOD", "COMPLIMENTARY BEVERAGE"]),
            ("H_COMP_OTHER", "Comp其他", ["COMP-ON PROPERTY", "COMP ON PROPERTY"]),
            ("H_LABOR", "人工成本", ["PAID TIME OFF", "13TH MONTH", "COMP LEAVE", "REWARD LEAVE",
                "OVERTIME", "PROVIDENT", "RETIREMENT FUND", "HEALTH & WELFARE", "HEALTH AND WELFARE",
                "WORKERS COMPENSATION", "EMPLOYEE HOUSING", "GROUP INSURANCE", "LIFE INSURANCE",
                "EMPLOYEE BENEFIT", "SOCIAL SECURITY", "BONUS", "OTHER ALLOWANCE"]),
            ("H_SPONSORSHIP", "贊助費", ["SPONSORSHIP", "SPONSOR", "CHARITABLE", "NON-PROFIT",
                "NONPROFIT", "捐贈", "公益"]),
            ("H_ADVERTISING", "廣告及推廣", ["PROMOTIONAL", "推廣", "PREPAID ADVERTISING",
                "ADVERTISING", "MEDIA-OTHER", "MEDIA - OTHER"]),
            ("H_MAINTENANCE", "維護費", ["R&M", "REPAIR", "MAINTENANCE", "MAINTEN", "UTILITIES",
                "UTILITY", "ELECTRICITY", "WATER USAGE", "GAS & OIL", "GAS AND OIL"]),
            ("H_EQUIP", "設施及器具採購", ["CHINAWARE", "CHINA/GLASS", "CHINA / GLASS", "GLASS/SILVER",
                "SILVERWARE", "KITCHEN SUPPL", "LINEN", "GAMBLING EQUIPMENT", "FACILITY INVENTORY",
                "GUEST SUPPLIES", "PANTRY SUPPLIES", "OPERATING ITEMS AND EQUIP", "LED WALL",
                "AV EQUIPMENT", "O/E-"]),
            ("H_PROFESSIONAL", "專業服務費", ["PROFESSIONAL FEE", "LEGAL FEE", "OUTSIDE SERVICE",
                "OUT SERV", "OUTSERV", "CONSULTING", "CONSULTANT"]),
        ]
        _assigned = pd.Series(False, index=combined.index)
        _summary = []
        for _hid, _lab, _toks in _RESCUE:
            _tokm = pd.Series(False, index=combined.index)
            for _t in _toks:
                _tokm |= _accU.str.contains(_re.escape(_t), regex=True, na=False)
            _m = _hO & ~_keepO & ~_assigned & _tokm
            if int(_m.sum()):
                combined.loc[_m, "horizontal_id"] = _hid
                if "horizontal_label" in combined.columns:
                    combined.loc[_m, "horizontal_label"] = _lab
                _assigned |= _m
                _amt = pd.to_numeric(combined.loc[_m, "amount_mop"], errors="coerce").abs().sum() / 1e4
                _summary.append(f"{_lab}={int(_m.sum()):,}行/{_amt:,.0f}萬")
        if _summary:
            print(f"  [其他救援→真H] " + " | ".join(_summary))

    # ── 內部資源 mark retag（user 2026-06-19 §A/§B）：項目組喺 databook mark 好 comp/staff，但我哋 H 錯位 ──
    #   comp_type → 對應 comp H（房落人工/票落Comp其他 等錯位修返）；is_labor → 人工（OVERRIDE capex 搶去建設）。
    #   tie 不變（H 唔喺 golden tie key；final_capex_opex 不動）。先 comp_type 再 is_labor（staff 最後贏）。
    if "horizontal_id" in combined.columns:
        # (a) comp_type → comp H
        if "comp_type" in combined.columns:
            _ct = combined["comp_type"].astype(str).str.strip()
            _CT2H = [
                ("H_HOTEL_ROOM", "Comp房間", ["客房支出", "房間支出", "Rooms", "Room", "Hotel Room",
                                              "Hotel room", "Hotel Misc", "Hotel-Spa"]),
                ("H_FNB", "Comp餐飲", ["F&B", "FnB", "FNB", "食品飲料支出", "食品與飲料支出"]),
                ("H_VENUE", "Comp活動場地", ["會場支出", "場地租借"]),
                ("H_COMP_TICKET", "Comp贈票", ["門票", "門票支出", "演唱會門票支出", "贈票支出"]),
                ("H_COMP_OTHER", "Comp其他", ["其他", "Others", "Other"]),
                ("H_OTHER", "其他", ["Flight"]),   # 機票=非內部資源運輸 → 其他
            ]
            _cnt = {}
            for _hid, _lab, _vals in _CT2H:
                _m = _ct.isin(_vals)
                if int(_m.sum()):
                    combined.loc[_m, "horizontal_id"] = _hid
                    if "horizontal_label" in combined.columns:
                        combined.loc[_m, "horizontal_label"] = _lab
                    _cnt[_lab] = int(_m.sum())
            if _cnt:
                print(f"  [comp_type retag] " + " | ".join(f"{k}={v:,}" for k, v in _cnt.items()))
        # (b) is_labor truthy → 人工（保留含 capex 勞工；HQ 對數時先取 opex 部分 — 見 recon_hq）
        if "is_labor" in combined.columns:
            _il = combined["is_labor"].astype(str).str.strip().str.lower()
            _ilm = ~_il.isin(["", "0", "0.0", "n", "no", "false", "f", "nan", "none", "<na>"])
            if int(_ilm.sum()):
                combined.loc[_ilm, "horizontal_id"] = "H_LABOR"
                if "horizontal_label" in combined.columns:
                    combined.loc[_ilm, "horizontal_label"] = "人工成本"
                _a = pd.to_numeric(combined.loc[_ilm, "amount_mop"], errors="coerce").abs().sum() / 1e4
                print(f"  [is_labor→人工] {int(_ilm.sum()):,} 行 / {_a:,.0f}萬")

    # ── wynn 人工成本 ⟺ Nature of Expenses 含 'Staff'（Staff Cost / Staff and Support Costs）──
    #   user 2026-06-19：wynn staff cost 規律 = account='Staff Cost'。非 Staff 嘅 人工 → 按性質改走。
    if "entity" in combined.columns and "account_desc" in combined.columns and "horizontal_id" in combined.columns:
        _wy = combined["entity"].astype(str).eq("wynn")
        _accw = combined["account_desc"].astype(str)
        _staff = _accw.str.contains("Staff", case=False, na=False)
        _hlw = combined["horizontal_id"].astype(str).str.strip().eq("H_LABOR")
        _to_lab = _wy & _staff & ~_hlw                       # Staff 但未係人工 → 補入
        _bad = _wy & _hlw & ~_staff                          # 非 Staff 但係人工 → 剷走
        _os = _bad & _accw.str.contains("Outside Service", case=False, na=False)
        _eqp = _bad & _accw.str.contains("Operating Items", case=False, na=False)
        _oth = _bad & ~_accw.str.contains("Outside Service|Operating Items", case=False, regex=True, na=False)
        for _m, _hid, _lab in [(_to_lab, "H_LABOR", "人工成本"), (_os, "H_PROFESSIONAL", "專業服務費"),
                               (_eqp, "H_EQUIP", "設施及器具採購"), (_oth, "H_OTHER", "其他")]:
            if int(_m.sum()):
                combined.loc[_m, "horizontal_id"] = _hid
                combined.loc[_m, "horizontal_label"] = _lab
        print(f"  [wynn 人工=Staff] 補入={int(_to_lab.sum())} | OutsideServ→專業={int(_os.sum())} | OpItems→設施={int(_eqp.sum())} | 其餘非Staff→其他={int(_oth.sum())}")

    # ── galaxy opex staff cost 規律（user 2026-06-19）：per-year allowlist（24/25 唔同）──
    # 24: code in {8000960,8000000,8000150}（8000140 Casual Labor 唔計 24；8107500 Event Service Fee
    #     項目組H 標 others 唔係 staff → 剔走，galaxy 24 由 +1,934 落 −620 ✓）
    # 25: code in {7926000,8000140,8000150,8000960} OR desc in 4 個（規律驗過 2025）
    if "entity" in combined.columns and "account_desc" in combined.columns and "account_code" in combined.columns and "horizontal_id" in combined.columns:
        _gl = combined["entity"].astype(str).eq("galaxy")
        _gopex = (~combined["final_capex_opex"].astype(str).str.strip().eq("Capex")
                  if "final_capex_opex" in combined.columns else pd.Series(True, index=combined.index))
        _gac = combined["account_code"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
        _gy2 = combined["year_bucket"].astype(str).str[:2]
        _g24 = _gy2.eq("24")
        _g25 = _gy2.eq("25")
        _ghl = combined["horizontal_id"].astype(str).str.strip().eq("H_LABOR")
        # 24 allowlist（code only；8107500 Event Service Fee 項目組H=others 已剔走）
        _gmatch24 = _gac.isin(["8000960", "8000000", "8000150"])
        # 25 allowlist（code OR desc）
        _gd25 = combined["account_desc"].astype(str).str.strip().isin(
            ["Casual Labor", "Contract Services", "Outsourced Contract Labor", "Payroll - Direct Event Investment"])
        _gmatch25 = _gd25 | _gac.isin(["7926000", "8000140", "8000150", "8000960"])
        # add: allowlist rows not yet H_LABOR（24 adds 8000000；25 adds event-specific accounts）
        _g_add24 = _gl & _gopex & _g24 & _gmatch24 & ~_ghl
        _g_add25 = _gl & _gopex & _g25 & _gmatch25 & ~_ghl
        # remove: H_LABOR rows outside respective allowlist
        _g_rm24  = _gl & _gopex & _g24 & _ghl & ~_gmatch24
        _g_rm25  = _gl & _gopex & _g25 & _ghl & ~_gmatch25
        _g_add = _g_add24 | _g_add25
        _g_rm  = _g_rm24  | _g_rm25
        if int(_g_add.sum()):
            combined.loc[_g_add, "horizontal_id"] = "H_LABOR"; combined.loc[_g_add, "horizontal_label"] = "人工成本"
        if int(_g_rm.sum()):
            combined.loc[_g_rm, "horizontal_id"] = "H_OTHER"; combined.loc[_g_rm, "horizontal_label"] = "其他"
        _aa = pd.to_numeric(combined.loc[_g_add, "amount_mop"], errors="coerce").abs().sum() / 1e4
        _ar = pd.to_numeric(combined.loc[_g_rm,  "amount_mop"], errors="coerce").abs().sum() / 1e4
        print(f"  [galaxy 人工=opex+acct規律] 補入={int(_g_add.sum())}行/{_aa:,.0f}萬 | 剷非命中人工→其他={int(_g_rm.sum())}行/{_ar:,.0f}萬")

    # ── comp 清理：comp-H 入面非內部 comp 嘅行搬走（user 2026-06-19 comp dump）──
    #   外部運輸/旅遊稅(非內部資源)→其他；Pre-Opening/Food Cost(運營,非comp)→其他；Operating Supplies/Items→設施
    if "horizontal_id" in combined.columns and "account_desc" in combined.columns:
        _COMPH = {"H_HOTEL_ROOM", "H_VENUE", "H_FNB", "H_COMP_TICKET", "H_COMP_OTHER"}
        _inc = combined["horizontal_id"].astype(str).str.strip().isin(_COMPH)
        _ad = combined["account_desc"].astype(str)
        _c_oth = _inc & _ad.str.contains("Comp External - Transp|Comp External - Tourism|Comp Transportation|External - Barter|Tourism Tax", case=False, regex=True, na=False)
        _c_po = _inc & _ad.str.contains("Pre-Opening|Pre Opening", case=False, regex=True, na=False)
        _c_fc = _inc & _ad.str.contains("Food Cost|Cost of Sales", case=False, regex=True, na=False)
        _c_os = _inc & _ad.str.contains("Operating Suppl|Operating Items", case=False, regex=True, na=False)
        _c_gt = _inc & _ad.str.contains("Government Tax", case=False, na=False)
        for _m, _hid, _lab in [(_c_oth, "H_OTHER", "其他"), (_c_po, "H_OTHER", "其他"),
                               (_c_fc, "H_OTHER", "其他"), (_c_os, "H_EQUIP", "設施及器具採購"),
                               (_c_gt, "H_OTHER", "其他")]:
            if int(_m.sum()):
                combined.loc[_m, "horizontal_id"] = _hid
                combined.loc[_m, "horizontal_label"] = _lab
        print(f"  [comp清理] 外部運輸/稅→其他={int(_c_oth.sum())} | Pre-Open→其他={int(_c_po.sum())} | FoodCost→其他={int(_c_fc.sum())} | Op供應→設施={int(_c_os.sum())} | 政府稅→其他={int(_c_gt.sum())}")

    # ── vml 25: comp account_code final-authority override（account_code 比 分類1 更可靠；25 bucket only，唔郁已 tie 嘅 24/23）──
    #   VML conf predominant_rules 已設呢啲係 comp accounts，但 row_horizontal_overrides.分類1 column_map 會蓋過佢。
    #   prep_tableau 最後再 enforce account_code：新 raw 有「Comp類型」但未 kedro re-run → 用 account_code 做 proxy。
    #   Future fix：VML conf row_horizontal_overrides 末加 {column_map: {col: "Comp類型",...}} → kedro re-run 後呢 block 可刪。
    if "entity" in combined.columns and "account_code" in combined.columns and "horizontal_id" in combined.columns:
        _vml25 = combined["entity"].astype(str).eq("vml") & combined["year_bucket"].astype(str).str[:2].eq("25")
        _vac = combined["account_code"].astype(str).str.strip()
        _vml_comp_ac = [
            (_vml25 & _vac.isin(["60001", "61001", "80541", "80544"]),
             "H_HOTEL_ROOM", "Comp房間"),
            (_vml25 & _vac.isin(["60002", "60003", "61002", "61003"]),
             "H_FNB", "Comp餐飲"),
            (_vml25 & _vac.isin(["60004", "60005"]),
             "H_COMP_TICKET", "Comp贈票"),
            (_vml25 & _vac.isin(["60007", "60010", "60099", "61007", "61010", "80185", "80410", "80580"]),
             "H_COMP_OTHER", "Comp其他"),
        ]
        _vcnt = {}
        for _m, _h, _l in _vml_comp_ac:
            n = int(_m.sum())
            if n:
                combined.loc[_m, "horizontal_id"] = _h
                combined.loc[_m, "horizontal_label"] = _l
                _vcnt[_l] = n
        print("  [vml 25 comp acct override] " + " | ".join(f"{k}={v}" for k, v in _vcnt.items()))

    # ── vml: COMPLIMENTARY OTHER → Comp活動場地（24 已咁 tag，23 冚晒落 comp其他；統一搬，修 23 會場 −1,024 + comp其他 +1,503）──
    if "entity" in combined.columns and "account_desc" in combined.columns and "horizontal_id" in combined.columns:
        _vco = (combined["entity"].astype(str) == "vml") \
            & combined["account_desc"].astype(str).str.contains("COMPLIMENTARY OTHER", case=False, na=False) \
            & combined["horizontal_id"].astype(str).str.strip().isin({"H_HOTEL_ROOM", "H_FNB", "H_COMP_TICKET", "H_COMP_OTHER"})
        if int(_vco.sum()):
            combined.loc[_vco, "horizontal_id"] = "H_VENUE"
            combined.loc[_vco, "horizontal_label"] = "Comp活動場地"
        print(f"  [vml COMPLIMENTARY OTHER→會場] {int(_vco.sum()):,} 行")

    # ── galaxy: 用項目組自己 H 標簽（項目組H = 基礎 ground truth）map comp/廣告/贊助/專業 ──
    #   格式：24-prefix = "Comp|Venue" (pipe)；25-prefix = "Comp-Venue" (dash)；大小寫不一（user 2026-06-21）。
    #   capex 不排除（項目組H comp capex 行也應計入，tie to pt_class_H filter）。
    #   cleanup after：galaxy y2=(24,25) 非 pt_class_H comp 行 → H_OTHER。
    if "entity" in combined.columns and "項目組H" in combined.columns and "horizontal_id" in combined.columns:
        _go = (combined["entity"].astype(str) == "galaxy")
        _ph = combined["項目組H"].astype(str)
        # split on first | or - ; case-insensitive "comp" check
        _one = _ph.str.extract(r'^([^|\-]+)')[0].str.strip().str.lower()
        # sub-category separator can be | or - ; match sub-type after separator
        _sub = lambda kw: _ph.str.contains(rf"[|\-].*(?:{kw})", case=False, regex=True, na=False)
        _gmap = [
            (_one.eq("comp") & _sub(r"room"),                                        "H_HOTEL_ROOM",   "Comp房間"),
            (_one.eq("comp") & _sub(r"venue"),                                       "H_VENUE",         "Comp活動場地"),
            (_one.eq("comp") & _sub(r"f&b|service charge"),                          "H_FNB",           "Comp餐飲"),
            (_one.eq("comp") & _sub(r"ticket|spa|printing fee"),                     "H_COMP_TICKET",   "Comp贈票"),
            (_one.eq("comp") & _sub(r"tax|travel|transport|credit card"),             "H_OTHER",         "其他"),
            (_one.eq("comp") & _sub(r"barter|others?|event service|gallery"),        "H_COMP_OTHER",    "Comp其他"),
            (_one.eq("marketing"),                                                    "H_ADVERTISING",   "廣告及推廣"),
            (_one.eq("sponsorship"),                                                  "H_SPONSORSHIP",   "贊助費"),
            (_ph.str.contains("professional fee", case=False, na=False),             "H_PROFESSIONAL",  "專業服務費"),
        ]
        _gcnt = []
        for _m, _h, _l in _gmap:
            _mm = _go & _m
            n = int(_mm.sum())
            _gcnt.append(f"{_l}={n}")
            if n:
                combined.loc[_mm, "horizontal_id"] = _h
                combined.loc[_mm, "horizontal_label"] = _l
        print("  [galaxy 項目組H→H] " + " | ".join(_gcnt))
        # cleanup: galaxy y2=(24,25) 非 pt_class_H comp → H_OTHER（只用 pt_class_H 做 comp ground truth）
        _COMPH_SET = {"H_HOTEL_ROOM", "H_VENUE", "H_FNB", "H_COMP_TICKET", "H_COMP_OTHER"}
        _y2g = combined["year_bucket"].astype(str).str[:2]
        _has_comp_ph = _one.eq("comp") & _go   # rows that WERE tagged by pt_class_H comp
        _gl_2425_comp = _go & _y2g.isin(["24", "25"]) & combined["horizontal_id"].astype(str).str.strip().isin(_COMPH_SET)
        _not_from_ph = _gl_2425_comp & ~_has_comp_ph
        _rm_n = int(_not_from_ph.sum())
        if _rm_n:
            combined.loc[_not_from_ph, "horizontal_id"] = "H_OTHER"
            combined.loc[_not_from_ph, "horizontal_label"] = "其他"
        print(f"  [galaxy 非pt_class_H comp剷走] {_rm_n:,}行（y2=24/25 comp但無pt_class_H→其他）")
        # galaxy 24-prefix Comp|Tax + Comp|Transport → H_COMP_OTHER（24 tie 需要；25 唔加 user 2026-06-21）
        _tt24 = (
            _go & _y2g.eq("24") & _one.eq("comp")
            & _ph.str.contains(r"[|\-].*(tax|transport)", case=False, regex=True, na=False)
            & combined["horizontal_id"].astype(str).str.strip().eq("H_OTHER")
        )
        if int(_tt24.sum()):
            combined.loc[_tt24, "horizontal_id"] = "H_COMP_OTHER"
            combined.loc[_tt24, "horizontal_label"] = "Comp其他"
        print(f"  [galaxy 24 Comp Tax/Transport→Comp其他] {int(_tt24.sum()):,}行")

    # ── melco/mgm: 從 人工 剷走 Contract/Travel/ProfFees/Uniform（spent-23/24 only；25 唔郁）user acct dump ──
    if "entity" in combined.columns and "account_desc" in combined.columns and "horizontal_id" in combined.columns:
        _2324 = combined["year_bucket"].astype(str).str[:2].isin(["23", "24"])
        _accx = combined["account_desc"].astype(str)
        _hlx = combined["horizontal_id"].astype(str).str.strip().eq("H_LABOR")
        _blob = (combined["account_code"].astype(str) + " " + _accx) if "account_code" in combined.columns else _accx
        _me = combined["entity"].astype(str).eq("melco") & _hlx & _2324
        _mg = combined["entity"].astype(str).eq("mgm") & _hlx & _2324
        _rules = [
            (_me & _blob.str.contains("611710|Contract Labo", case=False, regex=True, na=False), "H_PROFESSIONAL", "專業服務費"),
            (_me & _blob.str.contains("624110|Travel & Ent", case=False, regex=True, na=False), "H_OTHER", "其他"),
            (_mg & _blob.str.contains("588020|Professional Fee", case=False, regex=True, na=False), "H_PROFESSIONAL", "專業服務費"),
            (_mg & _blob.str.contains("590010|Travel & Acc|Travel & Ent", case=False, regex=True, na=False), "H_OTHER", "其他"),
            (_mg & _blob.str.contains("578000|Uniform", case=False, regex=True, na=False), "H_OTHER", "其他"),
        ]
        _cnt = 0
        for _m, _hid, _lab in _rules:
            if int(_m.sum()):
                combined.loc[_m, "horizontal_id"] = _hid
                combined.loc[_m, "horizontal_label"] = _lab
                _cnt += int(_m.sum())
        print(f"  [melco/mgm 人工剷 spent-23/24] {_cnt:,} 行")

    # ── vml: 部門間人工分攤撥出（89001 INTERDEPARTMENT ALLOC-OTHER）唔係真正人工成本 → H_OTHER ──
    if "entity" in combined.columns and "account_code" in combined.columns and "horizontal_id" in combined.columns:
        _vml_alloc = (
            combined["entity"].astype(str).eq("vml")
            & combined["account_code"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True).eq("89001")
            & combined["horizontal_id"].astype(str).str.strip().eq("H_LABOR")
        )
        n = int(_vml_alloc.sum())
        if n:
            combined.loc[_vml_alloc, "horizontal_id"] = "H_OTHER"
            combined.loc[_vml_alloc, "horizontal_label"] = "其他"
        _va = pd.to_numeric(combined.loc[_vml_alloc, "amount_mop"], errors="coerce").sum() / 1e4
        print(f"  [vml 89001 部門間分攤→其他] {n}行/{_va:,.0f}萬")

    # ── mgm: NG0 博彩 opex payroll 移出 H_LABOR（博彩營運人工 ≠ investment staff；user 2026-06-22 確認）──
    #   只 MGM NG0 有 gaming opex payroll（其餘5家=0）；24 opex staff 10,193−3,152=7,041≈HQ6,829(Δ+212)。
    #   只郁 opex（capex tie 維持 −5）；全年適用（principle year-agnostic，跑完核 25(+871)/23(+639) 唔好爆）。
    if "entity" in combined.columns and "ng_code" in combined.columns and "horizontal_id" in combined.columns:
        _mg_ng0 = (
            combined["entity"].astype(str).eq("mgm")
            & combined["ng_code"].astype(str).str.strip().str.upper().eq("NG0")
            & combined["horizontal_id"].astype(str).str.strip().eq("H_LABOR")
        )
        if "final_capex_opex" in combined.columns:
            _mg_ng0 &= ~combined["final_capex_opex"].astype(str).str.strip().eq("Capex")
        _nn = int(_mg_ng0.sum())
        if _nn:
            combined.loc[_mg_ng0, "horizontal_id"] = "H_OTHER"
            combined.loc[_mg_ng0, "horizontal_label"] = "其他"
        _aa = pd.to_numeric(combined.loc[_mg_ng0, "amount_mop"], errors="coerce").abs().sum() / 1e4
        print(f"  [mgm NG0博彩人工→其他] {_nn}行/{_aa:,.0f}萬（博彩營運人工非investment staff）")

    # ── capex staff 補返人工（項目組H mark 咗 staff 但 capex enforcement 搶咗去建設；user Table 1 capex+opex）──
    if "項目組H" in combined.columns and "final_capex_opex" in combined.columns and "horizontal_id" in combined.columns:
        _capstaff = combined["final_capex_opex"].astype(str).str.strip().eq("Capex") & \
            combined["項目組H"].astype(str).str.contains("staff cost|Staff cost|人工成本", case=False, regex=True, na=False) & \
            ~combined["horizontal_id"].astype(str).str.strip().eq("H_LABOR")
        if int(_capstaff.sum()):
            combined.loc[_capstaff, "horizontal_id"] = "H_LABOR"
            combined.loc[_capstaff, "horizontal_label"] = "人工成本"
            _a = pd.to_numeric(combined.loc[_capstaff, "amount_mop"], errors="coerce").abs().sum() / 1e4
            print(f"  [capex staff→人工] {int(_capstaff.sum()):,} 行 / {_a:,.0f}萬（項目組H=staff 嘅 capex）")

    # ── 100% fill：4 欄唔好有 blank/None（user 要全部填滿）──
    #   project 名空 → dicj code；subproject code 空 → dicj code；subproject 名空 → project(DICJ名)
    def _fill100(dst, src):
        if dst in combined.columns and src in combined.columns:
            _d = combined[dst].astype(str).str.strip()
            _b = _d.isin(["", "nan", "None", "NaN", "<NA>"])
            if int(_b.sum()):
                combined.loc[_b, dst] = combined.loc[_b, src].astype(str).str.strip()
                print(f"  [fill100] {dst}: 補 {int(_b.sum()):,} 個空 ← {src}")
    _fill100("project", "dicj code")
    _fill100("subproject code", "dicj code")
    _fill100("subproject", "project")
    # subproject 名去 code 前綴（code 已有獨立欄，名唔使重複）：vml 'SP00033 Comprehensive'→'Comprehensive'
    if "subproject" in combined.columns and "subproject code" in combined.columns:
        def _strip_pre(n, c):
            n = str(n).strip(); c = str(c).strip()
            if c and n.startswith(c):
                _r = n[len(c):].lstrip(" -:_.．、|/")
                return _r if _r else n
            return n
        _before = sum(1 for n, c in zip(combined["subproject"].tolist(), combined["subproject code"].tolist())
                      if str(c).strip() and str(n).strip().startswith(str(c).strip()))
        combined["subproject"] = [_strip_pre(n, c) for n, c in
                                  zip(combined["subproject"].tolist(), combined["subproject code"].tolist())]
        print(f"  [subproject名去碼前綴] {_before:,} 行")
    for _c in ("dicj code", "project", "subproject code", "subproject"):
        if _c in combined.columns:
            _nb = int(combined[_c].astype(str).str.strip().isin(["", "nan", "None", "NaN", "<NA>"]).sum())
            print(f"  [4層 100%?] {_c}: blank={_nb}")

    # ── blank V 填 NG 主 V（user: blank V 也是問題）。只填空，唔郁已有 V。NG primary = ng_default canonical。──
    NG_PRIMARY = {"NG0": ("V_GAMING_VENUE", "博彩娛樂場優化"), "NG1": ("V_OVERSEAS_ROADSHOW", "參與海外路演與宣傳活動"),
                  "NG2": ("V_MICE", "會議展覽"), "NG3": ("V_CONCERT", "娛樂表演"), "NG4": ("V_SPORT_EVENT", "體育賽事"),
                  "NG5": ("V_ART_EXHIBITION", "文藝展覽表演"), "NG6": ("V_WELLNESS", "健康養生"), "NG7": ("V_THEME_PARK", "主題遊樂場地"),
                  "NG8": ("V_RESTAURANT", "餐廳"), "NG9": ("V_COMMUNITY", "社區活化"), "NG10": ("V_MARITIME", "海上旅遊"),
                  "NG11": ("V_OTHER", "其他")}
    # 透明度 flag v_調整：""=原始真V / "填補"=原blank填NG主V / "重設"=掛錯枝reset。Tableau 可 filter 走睇原始真 V。
    combined["v_調整"] = ""
    _ngc = combined["ng_code"].astype(str).str.strip() if "ng_code" in combined.columns else pd.Series("", index=combined.index)
    # (1) blank V → 填 NG 主 V
    if "vertical_label" in combined.columns and "ng_code" in combined.columns:
        _vlb = combined["vertical_label"].astype(str).str.strip().isin(["", "nan", "None", "NaN", "<NA>"])
        if int(_vlb.sum()):
            _by_ng0 = int((_vlb & _ngc.eq("NG0")).sum())
            print(f"  [blank V] {int(_vlb.sum()):,} 行冇 V（其中 NG0 博彩 {_by_ng0:,}）→ 填 NG 主 V")
            combined.loc[_vlb, "v_調整"] = "填補"
            combined.loc[_vlb, "vertical_label"] = _ngc[_vlb].map(lambda n: NG_PRIMARY.get(n, ("", "其他"))[1])
            if "vertical_id" in combined.columns:
                _vib = combined["vertical_id"].astype(str).str.strip().isin(["", "nan", "None", "NaN", "<NA>"])
                combined.loc[_vib, "vertical_id"] = _ngc[_vib].map(lambda n: NG_PRIMARY.get(n, ("V_OTHER",))[0])
    # (2) V 掛錯枝 enforcement：V 唔喺該 NG eligible 集（且唔係 cross-cutting）→ reset NG 主 V。
    #     cross-cutting（任何 NG 合理，永不 reset）= 物業優化 + 公共設施 + 社區活化（user 2026-06-17 確認；拍攝視頻 reset）。
    _ngcats = cats.get("ng_categories") or {}
    CROSS_CUTTING = {"V_PROPERTY_UPGRADE", "V_PUBLIC_FACILITY", "V_COMMUNITY"}
    if _ngcats and "ng_code" in combined.columns and "vertical_id" in combined.columns:
        _elig = {ng: (set(d.get("eligible_verticals") or []) | CROSS_CUTTING | {"V_OTHER"})
                 for ng, d in _ngcats.items()}
        _vid3 = combined["vertical_id"].astype(str).str.strip()
        _f3 = combined["v_調整"].astype(str).str.strip()
        _stray = pd.Series([bool(v) and fl == "" and (n in _elig) and (v not in _elig[n])
                            for v, n, fl in zip(_vid3.tolist(), _ngc.tolist(), _f3.tolist())],
                           index=combined.index)
        if int(_stray.sum()):
            _samt = pd.to_numeric(combined.loc[_stray, "amount_mop"], errors="coerce").abs().sum() / 1e6
            print(f"  [V掛錯枝] reset {int(_stray.sum()):,} 行 V→NG主V（Σ={_samt:,.0f}M；保留設施/社區 cross-cutting）")
            combined.loc[_stray, "v_調整"] = "重設"
            combined.loc[_stray, "vertical_label"] = _ngc[_stray].map(lambda n: NG_PRIMARY.get(n, ("", "其他"))[1])
            combined.loc[_stray, "vertical_id"] = _ngc[_stray].map(lambda n: NG_PRIMARY.get(n, ("V_OTHER",))[0])

    # ── 點名 NG×V 修正（user §0：V 對唔上 databook NG → align 去 NG-native；如刻意 two-taxonomy 可 raw 改 NG）──
    _FIXNGV = [  # (entity, ng_code, project 包含, → vertical_id)
        ("sjm", "NG5", "澳門設計大獎", "V_ART_EXHIBITION"),
        ("galaxy", "NG9", "為前線團隊成員提供年度培訓", "V_COMMUNITY"),
        ("vml", "NG9", "帶動旅客進入社區", "V_COMMUNITY"),
        ("melco", "NG9", "新濠風尚", "V_COMMUNITY"),
    ]
    if "vertical_id" in combined.columns and "project" in combined.columns and "entity" in combined.columns:
        _proj = combined["project"].astype(str)
        _ent = combined["entity"].astype(str)
        for _e, _n, _pk, _vid in _FIXNGV:
            _m = _ent.eq(_e) & _ngc.eq(_n) & _proj.str.contains(_pk, na=False, regex=False)
            if int(_m.sum()):
                combined.loc[_m, "vertical_id"] = _vid
                print(f"  [NGV修正] {_e} {_n} {_pk[:12]} → {_vid}（{int(_m.sum())}行）")

    # ── (3) V label 大改（user 2026-06-18）：更名/合併/NG條件 = 確定；拆分 = keyword heuristic（待 dump 驗）──
    if "vertical_id" in combined.columns and "vertical_label" in combined.columns:
        _vid = combined["vertical_id"].astype(str).str.strip().tolist()
        _ngl = _ngc.tolist()
        _basis = (combined.get("account_desc", pd.Series("", index=combined.index)).astype(str) + " " +
                  combined.get("description", pd.Series("", index=combined.index)).astype(str) + " " +
                  combined.get("project", pd.Series("", index=combined.index)).astype(str) + " " +
                  combined.get("項目組V", pd.Series("", index=combined.index)).astype(str)).tolist()
        _cur = combined["vertical_label"].astype(str).tolist()

        def _has(s, kws):
            return any(k in s for k in kws)

        def _relabel(v, ng, s, cur):
            # ── 確定性 ──
            if v == "V_MICE": return "會展活動"
            if v == "V_WELLNESS": return "康養活動"
            if v == "V_MARITIME": return "海上活動"
            if v in ("V_VENUE_PERF_SPORT_MICE", "V_MUSEUM"): return "內部設施-場館"  # 場館+博物館 併入場館
            if v == "V_RESTAURANT": return "內部設施-餐廳"
            if v == "V_THEME_PARK": return "內部設施-遊樂"
            if v in ("V_OVERSEAS_OFFICE", "V_REGIONAL_TEAM"): return "國際代表團隊及海外辦事處"
            if v == "V_OVERSEAS_WEB_SEO": return "宣傳推廣"
            if v == "V_PROPERTY_UPGRADE":
                if ng == "NG2": return "內部設施-場館"   # 會展場館優化 → 場館（user 2026-06-18）
                if ng == "NG6": return "內部設施-康養"
                if ng == "NG11": return "內部設施-其他"
                return "內部設施優化"
            if v == "V_PUBLIC_FACILITY": return "外部設施優化"
            # ── 宣傳：項目組覺得線上/線下難分 → 改回「路演」+「宣傳推廣」兩類（user 2026-06-18）──
            #   路演 = 參加路演(V_OVERSEAS_ROADSHOW，贊助/政府嗰啲歸 政府公益社區活動)；
            #   其餘宣傳（邀請旅行社[含佣金]/邀請外國客/拍片/海外SEO）= 宣傳推廣。
            if v == "V_OVERSEAS_ROADSHOW":
                return "政府、公益及社區活動" if _has(s, ("贊助", "政府", "MGTO", "旅遊局", "文化局", "體育局", "運動會", "Sponsor", "sponsor")) else "路演"
            if v in ("V_INVITE_AGENCY", "V_INVITE_GUEST", "V_PROMO_VIDEO"): return "宣傳推廣"
            if v == "V_CONCERT":
                return "節日慶典" if _has(s, ("節", "慶典", "賀歲", "新年", "CNY", "Festival", "花車", "煙花", "巡遊", "花燈", "Parade", "parade")) else "演出表演"
            if v == "V_FOOD_EVENT":
                if _has(s, ("Deposits", "CIP", "資產", "Renovation", "Asset Under", "新工作范圍", "工程", "建設", "裝修")): return "內部設施-餐廳"
                if _has(s, ("宴", "Dinner", "dinner", "菜單", "Menu", "名廚", "Chef", "chef", "美酒", "Wine", "wine", "Pairing", "Whisky")): return "特別菜單或宴會"
                if _has(s, ("媒體", "Media", "media", "推廣", "KOL", "宣傳", "Promo")): return "宣傳推廣"
                return "美食-其他"
            if v == "V_COMMUNITY":
                # 社區活化拆 3：建設類分內/外（公共片區=外）；其餘=政府公益（user 2026-06-18 §2 sjm 新馬路）
                _con = _has(s, ("建設", "工程", "設施", "裝修", "Construction", "Renovation", "renovation",
                                "重整", "碼頭", "環境美化", "Deposits paid", "CIP", "活化計劃"))
                if _con:
                    return "外部設施-社區活化" if _has(s, ("新馬路", "片區", "碼頭", "福隆新街", "公共", "環境美化")) else "內部設施-社區活化"
                return "政府、公益及社區活動"
            return cur
        _newlab = [_relabel(v, n, s, c) for v, n, s, c in zip(_vid, _ngl, _basis, _cur)]
        _nch = sum(1 for a, b in zip(_newlab, _cur) if a != b)
        combined["vertical_label"] = _newlab
        from collections import Counter
        _top = Counter(_newlab).most_common(14)
        print(f"  [V relabel] 改 {_nch:,} 行；新 V top: " + " | ".join(f"{k}={v}" for k, v in _top))

    # ── NG1 路演 清理（user 2026-06-22：好多唔似路演；V 廣播/reset 令水上樂園/演唱會 冚入路演）──
    #   路演/展銷 keyword→留路演；推廣/營銷/媒體/品牌/航空/銷售/辦事處 等→宣傳推廣；其餘(冇信心)→其他。
    if "ng_code" in combined.columns and "vertical_label" in combined.columns:
        _n1r = _ngc.eq("NG1") & combined["vertical_label"].astype(str).str.strip().eq("路演")
        _nmx = (combined.get("project", pd.Series("", index=combined.index)).astype(str) + " "
                + combined.get("subproject", pd.Series("", index=combined.index)).astype(str))
        _road = _nmx.str.contains(r"路演|展銷|展销|銷售路演|road\s*show", case=False, regex=True, na=False)
        _promo = _nmx.str.contains(
            r"推廣|推广|營銷|营销|媒體|媒体|廣告|广告|宣傳|宣传|品牌|marketing|航班|航線|航线|航空|"
            r"辦事處|办事处|代表團|代表团|銷售|销售|客源|旅行社|OTA|網站|网站|市場|市场|代理",
            case=False, regex=True, na=False)
        _toPromo = _n1r & ~_road & _promo
        _toOther = _n1r & ~_road & ~_promo
        if int(_toPromo.sum()):
            combined.loc[_toPromo, "vertical_label"] = "宣傳推廣"
            if "vertical_id" in combined.columns:
                combined.loc[_toPromo, "vertical_id"] = "V_OVERSEAS_WEB_SEO"
        if int(_toOther.sum()):
            combined.loc[_toOther, "vertical_label"] = "其他"
            if "vertical_id" in combined.columns:
                combined.loc[_toOther, "vertical_id"] = "V_OTHER"
        print(f"  [NG1路演清理] 留路演={int((_n1r & _road).sum())} 宣傳推廣={int(_toPromo.sum())} 其他={int(_toOther.sum())}")

    # ── NG6 康養：設施/中心/診所/水療/會所 等 → 內部設施-康養（user 2026-06-18 §4：康養設施較細）──
    #   康養活動(活動/瑜伽節/講座) 維持；facility kw（含 subproject）先升做 內部設施-康養。
    if "ng_code" in combined.columns and "vertical_label" in combined.columns:
        _ng6 = _ngc.eq("NG6") & combined["vertical_label"].astype(str).eq("康養活動")
        _fb6 = (combined.get("project", pd.Series("", index=combined.index)).astype(str) + " " +
                combined.get("subproject", pd.Series("", index=combined.index)).astype(str) + " " +
                combined.get("description", pd.Series("", index=combined.index)).astype(str) + " " +
                combined.get("account_desc", pd.Series("", index=combined.index)).astype(str))
        _FAC6 = ("中心", "診所", "會所", "水療", "理療", "中醫房", "醫療", "保健", "設施", "基礎設施",
                 "SPA", "Spa", "spa", "Clinic", "Polyclinic", "Surgical", "Wellness Cent",
                 "Wellness Centre", "Wellness Center", "Fitness", "Gym")
        _fac6 = _ng6 & _fb6.apply(lambda s: any(k in s for k in _FAC6))
        if int(_fac6.sum()):
            combined.loc[_fac6, "vertical_label"] = "內部設施-康養"
            if "vertical_id" in combined.columns:
                combined.loc[_fac6, "vertical_id"] = "V_PROPERTY_UPGRADE"
            _a6 = pd.to_numeric(combined.loc[_fac6, "amount_mop"], errors="coerce").abs().sum() / 1e4
            print(f"  [NG6 康養設施→內部設施-康養] {int(_fac6.sum()):,} 行 / {_a6:,.0f}萬")

    # ── sjm 24: 食飲 comp 入面 演出/體育/會展/節慶 V 其實係 event/venue comp（ct=FnB 分唔到 venue，用 V 分）──
    #   ⚠ 必須喺 V relabel 之後（vertical_label 先係最終值）。golden 會場 3,938 ≈ 演出表演+體育賽事+會展活動+節日慶典
    if "entity" in combined.columns and "vertical_label" in combined.columns and "horizontal_id" in combined.columns:
        _sjv = (combined["entity"].astype(str) == "sjm") & (combined["year_bucket"].astype(str) == "24") \
            & combined["horizontal_id"].astype(str).str.strip().eq("H_FNB") \
            & combined["vertical_label"].astype(str).str.strip().isin({"演出表演", "體育賽事", "會展活動", "節日慶典"})
        if int(_sjv.sum()):
            combined.loc[_sjv, "horizontal_id"] = "H_VENUE"
            combined.loc[_sjv, "horizontal_label"] = "Comp活動場地"
        print(f"  [sjm 24 食飲→會場 by V] {int(_sjv.sum()):,} 行")

    # ── H reasonableness 修正（2026-06-20 audit 紅旗）：account_desc 對唔上 label 嘅搬返正確 H。tie-safe（淨 re-tag）──
    if "horizontal_id" in combined.columns and "account_desc" in combined.columns and "entity" in combined.columns:
        _hid = combined["horizontal_id"].astype(str).str.strip()
        _ad = combined["account_desc"].astype(str).str.strip()
        _ent = combined["entity"].astype(str)
        # 項目組H=staff 豁免：項目組親標 staff 嘅 capex 人工（galaxy AUC-Design&Dev）唔好被 人工→建設 搬走
        # （user 2026-06-22：呢啲係真·資本化人工 hint，capex 內部保留做 H_LABOR；只 galaxy 命中，wynn 項目組H空不受影響）
        _pthx = combined.get("項目組H", pd.Series("", index=combined.index)).astype(str).str.contains(
            r"staff\s*cost|人工成本", case=False, regex=True, na=False)
        _fixes = [
            (_hid.eq("H_LICENSE") & _ad.str.contains("CONTRACT ENTERTAINMENT", case=False, na=False), "H_PERFORMER", "藝人演出費", "授權→藝人"),
            (_hid.eq("H_ADVERTISING") & _ad.str.contains("Contract Performers|Performers & Contract", case=False, regex=True, na=False), "H_PERFORMER", "藝人演出費", "廣告→藝人"),
            (_hid.eq("H_LABOR") & ~_pthx & _ad.str.contains(r"Renovation|\bCIP\b|\bAUC\b|Building Improvement|Capital Project", case=False, regex=True, na=False), "H_CONSTRUCTION", "建設與設施支出", "人工→建設"),
            (_hid.eq("H_OTHER") & _ad.str.lower().eq("salaries") & _ent.ne("galaxy"), "H_LABOR", "人工成本", "其他→人工(非gx)"),
            (_hid.eq("H_PERFORMER") & _ad.str.contains("Repairs.{0,3}Maintenance|Utilities", case=False, regex=True, na=False), "H_MAINTENANCE", "維護費", "藝人→維護"),
            (_hid.eq("H_FNB") & _ad.str.contains("FLOWERS", case=False, na=False), "H_OTHER", "其他", "餐飲→其他(花)"),
            (_hid.eq("H_FNB") & _ad.str.contains("Promotional chip", case=False, na=False), "H_ADVERTISING", "廣告及推廣", "餐飲→廣告(籌碼)"),
            (_hid.eq("H_COMP_OTHER") & _ad.str.contains("Apple Product", case=False, na=False), "H_EQUIP", "設施及器具採購", "comp其他→設施(Apple)"),
            (_hid.eq("H_COMP_OTHER") & _ad.str.contains(r"TRAVEL & ENT TRANSP|TRAVEL CUSTOMER|Comp External - Transp|^Transportation$", case=False, regex=True, na=False), "H_OTHER", "其他", "comp其他→其他(運輸)"),
        ]
        _msgs = []
        for _m, _nid, _nlab, _tag in _fixes:
            n = int(_m.sum())
            _msgs.append(f"{_tag}={n}")
            if n:
                combined.loc[_m, "horizontal_id"] = _nid
                combined.loc[_m, "horizontal_label"] = _nlab
        print("  [H紅旗修] " + " | ".join(_msgs))

    # ── 全6家：capex is_labor non-blank-non-0 → H_LABOR（user 2026-06-22：每家都靠 is_labor 識別 capex 人工）──
    #   放 H紅旗修 之後（免 AUC/CIP/Renovation 被 人工→建設 undo）。**只 capex**（opex staff 已 tie，唔好掂；
    #   capex 內部重分類唔郁 capex total）。捉返各家資本化人工，改善 Table 1。
    if "is_labor" in combined.columns and "horizontal_id" in combined.columns and "final_capex_opex" in combined.columns:
        _il = combined["is_labor"].astype(str).str.strip()
        _il_on = ~_il.isin(["", "0", "0.0", "nan", "None", "NaN", "N", "n", "False", "false", "FALSE", "No", "no"])
        _cx = combined["final_capex_opex"].astype(str).str.strip().eq("Capex")
        _cap_lab = (_il_on & _cx & ~combined["horizontal_id"].astype(str).str.strip().eq("H_LABOR"))
        _n = int(_cap_lab.sum())
        if _n:
            combined.loc[_cap_lab, "horizontal_id"] = "H_LABOR"
            combined.loc[_cap_lab, "horizontal_label"] = "人工成本"
        # per-entity 量
        _byent = {}
        if _n:
            _tmp = combined.loc[_cap_lab].copy()
            _tmp["_a"] = pd.to_numeric(_tmp["amount_mop"], errors="coerce").abs()
            for _e, _v in _tmp.groupby(_tmp["entity"].astype(str))["_a"].sum().items():
                _byent[_e] = _v / 1e4
        print(f"  [capex is_labor→人工(全家)] {_n}行 | " + " ".join(f"{k}={v:,.0f}萬" for k, v in _byent.items()))

    # ── melco 23: intercompany labor / payroll allocation 剷出人工（user golden 23=4,172；allocation≠直接人工）──
    #   只 melco 23（24 已 −388 唔郁；25 唔郁）。841000/841005 Interco Labor Charge、611199 Allocation A/C-Payroll → 其他。
    if "entity" in combined.columns and "account_code" in combined.columns and "horizontal_id" in combined.columns:
        _meblob = (combined["account_code"].astype(str) + " "
                   + combined.get("account_desc", pd.Series("", index=combined.index)).astype(str))
        _me23 = (combined["entity"].astype(str).eq("melco")
                 & combined["year_bucket"].astype(str).eq("23")
                 & combined["horizontal_id"].astype(str).str.strip().eq("H_LABOR")
                 & _meblob.str.contains(r"841000|841005|611199|Interco Labor Charge Alloc|Allocation A/C - Payroll",
                                        case=False, regex=True, na=False))
        _n2 = int(_me23.sum())
        if _n2:
            combined.loc[_me23, "horizontal_id"] = "H_OTHER"
            combined.loc[_me23, "horizontal_label"] = "其他"
        _a2 = pd.to_numeric(combined.loc[_me23, "amount_mop"], errors="coerce").sum() / 1e4
        print(f"  [melco 23 allocation剷人工] {_n2}行/{_a2:,.0f}萬")

    # ── V 紅旗修（2026-06-20 audit）──
    # B. 藝術珍寶博物館（mgm）統一 → 內部設施-場館（user：博物館=場館；之前一半派咗 文藝展覽表演）
    if "vertical_label" in combined.columns and "subproject" in combined.columns:
        _mus = combined["subproject"].astype(str).str.contains("藝術珍寶博物館|珍寶博物館", na=False) \
            & combined["vertical_label"].astype(str).str.strip().eq("文藝展覽表演")
        if int(_mus.sum()):
            combined.loc[_mus, "vertical_id"] = "V_MUSEUM"
            combined.loc[_mus, "vertical_label"] = "內部設施-場館"
        print(f"  [V紅旗修] 藝術珍寶博物館→內部設施-場館 {int(_mus.sum())}行")
    # A. sjm subproject 垃圾名（附件3.pdf 等）→ 用 project 真名
    if "entity" in combined.columns and "subproject" in combined.columns and "project" in combined.columns:
        _sp = combined["subproject"].astype(str)
        _bad = (combined["entity"].astype(str) == "sjm") \
            & (_sp.str.contains("附件", na=False) | _sp.str.lower().str.endswith(".pdf", na=False)) \
            & combined["project"].astype(str).str.strip().ne("")
        if int(_bad.sum()):
            combined.loc[_bad, "subproject"] = combined.loc[_bad, "project"]
        print(f"  [sjm 垃圾名→project] {int(_bad.sum())}行")

    # C. galaxy: 描述含「Comp Rooms -」（房 service/tax）而家喺 餐飲/其他 → 搬返 Comp房間（JE 內 re-tag，tie-safe）
    if "entity" in combined.columns and "description" in combined.columns and "horizontal_id" in combined.columns:
        _cr = (combined["entity"].astype(str) == "galaxy") \
            & combined["description"].astype(str).str.contains(r"Comp Room", case=False, regex=True, na=False) \
            & ~combined["horizontal_id"].astype(str).str.strip().eq("H_HOTEL_ROOM")
        if int(_cr.sum()):
            combined.loc[_cr, "horizontal_id"] = "H_HOTEL_ROOM"
            combined.loc[_cr, "horizontal_label"] = "Comp房間"
        print(f"  [galaxy Comp Rooms desc→房] {int(_cr.sum())}行")

    # D. sjm: comp 入面 account_desc=「專業服務費」其實係專業（sjm 24 comp over）→ 搬返 H_PROFESSIONAL
    if "entity" in combined.columns and "account_desc" in combined.columns and "horizontal_id" in combined.columns:
        _COMPH2 = {"H_HOTEL_ROOM", "H_VENUE", "H_FNB", "H_COMP_TICKET", "H_COMP_OTHER"}
        _sjp = (combined["entity"].astype(str) == "sjm") \
            & combined["account_desc"].astype(str).str.strip().eq("專業服務費") \
            & combined["horizontal_id"].astype(str).str.strip().isin(_COMPH2)
        if int(_sjp.sum()):
            combined.loc[_sjp, "horizontal_id"] = "H_PROFESSIONAL"
            combined.loc[_sjp, "horizontal_label"] = "專業服務費"
        print(f"  [sjm comp 專業服務費→專業] {int(_sjp.sum())}行")

    # E. sjm: comp 入面非 comp 嘅雜項 → H_OTHER（購貨-食品/Supplies/Low value = 費用，唔係贈品）
    if "entity" in combined.columns and "account_desc" in combined.columns and "horizontal_id" in combined.columns:
        _COMPH3 = {"H_HOTEL_ROOM", "H_VENUE", "H_FNB", "H_COMP_TICKET", "H_COMP_OTHER"}
        _sj_ent = combined["entity"].astype(str).eq("sjm")
        _sj_comp = combined["horizontal_id"].astype(str).str.strip().isin(_COMPH3)
        _sj_ad = combined["account_desc"].astype(str).str.strip()
        _bad_comp = _sj_ent & _sj_comp & (
            _sj_ad.str.contains(r"購貨\s*[-－]?\s*食品|Supplies.{0,15}Sanitary|Low value purchase|Low Value Purchase", regex=True, na=False)
            | _sj_ad.str.contains("各項招待費", na=False)
        )
        if int(_bad_comp.sum()):
            combined.loc[_bad_comp, "horizontal_id"] = "H_OTHER"
            combined.loc[_bad_comp, "horizontal_label"] = "其他"
        print(f"  [sjm comp 非贈品→其他] {int(_bad_comp.sum())}行（購貨食品/Supplies/Low value/招待費）")

    # F. sjm 25 H_COMP_OTHER blank comp_type → H_ADVERTISING
    # 25 raw 全行 comp_type 空白，H_COMP_OTHER 推廣費 ~747萬係 sjm25 over +729 嘅來源；
    # 24/23 comp_type 有值（其他/禮品及禮劵）唔受此 condition 影響
    if "entity" in combined.columns and "horizontal_id" in combined.columns:
        _sjm25_mask = (
            combined["entity"].astype(str).eq("sjm")
            & combined["year_bucket"].astype(str).str[:2].eq("25")
            & combined["horizontal_id"].astype(str).str.strip().eq("H_COMP_OTHER")
            & combined.get("comp_type", pd.Series("", index=combined.index)).astype(str).str.strip().isin(["", "nan", "None", "NaN"])
        )
        if int(_sjm25_mask.sum()):
            combined.loc[_sjm25_mask, "horizontal_id"] = "H_ADVERTISING"
            combined.loc[_sjm25_mask, "horizontal_label"] = "廣告及推廣"
        print(f"  [sjm 25 comp其他blank→廣告] {int(_sjm25_mask.sum())}行")

    # G. mgm: comp 入面非 internal resource → 正確 H
    # 注意：MGM comp_type 全部 blank，所有分類靠 account_desc + description
    # Airfare/Ferry = 運輸; Outside Comp = 外部comp; Other Electronic = 器具
    # OC Trans Upfront Offer = 保留（HQ 計入 comp，移走會 over-correct 24 by ~1,083萬）
    # Travel Subsidy (description) = 海外旅費補貼現金（非 internal resource）→ H_OTHER
    if "entity" in combined.columns and "horizontal_id" in combined.columns:
        _COMPH_G = {"H_HOTEL_ROOM", "H_VENUE", "H_FNB", "H_COMP_TICKET", "H_COMP_OTHER"}
        _mg = combined["entity"].astype(str).eq("mgm")
        _mg_comp = combined["horizontal_id"].astype(str).str.strip().isin(_COMPH_G)
        _mg_ad = combined.get("account_desc", pd.Series("", index=combined.index)).astype(str).str.strip()
        _mg_ds = combined.get("description", pd.Series("", index=combined.index)).astype(str).str.strip()
        _mgm_fixes_g = [
            (_mg & _mg_comp & _mg_ad.eq("Airfare"),
             "H_OTHER", "其他", "Airfare→其他"),
            (_mg & _mg_comp & _mg_ad.str.contains("Outside Comp", na=False),
             "H_OTHER", "其他", "外部Comp→其他"),
            (_mg & _mg_comp & _mg_ad.eq("Other Electronic"),
             "H_EQUIP", "設施及器具採購", "電子器具→設施"),
            (_mg & _mg_comp & _mg_ad.str.contains(r"^Ferry", na=False),
             "H_OTHER", "其他", "Ferry→其他"),
            (_mg & _mg_comp & _mg_ds.str.contains(r"Travel Subsid|International Travel Subsid|Promotional Cash for Overseas", case=False, regex=True, na=False),
             "H_OTHER", "其他", "海外旅費補貼→其他"),
        ]
        _gmsgs = []
        for _m, _hid, _lab, _tag in _mgm_fixes_g:
            n = int(_m.sum())
            _gmsgs.append(f"{_tag}={n}")
            if n:
                combined.loc[_m, "horizontal_id"] = _hid
                combined.loc[_m, "horizontal_label"] = _lab
        print("  [mgm comp非internal] " + " | ".join(_gmsgs))

    # ── tie-safe 剔走全 0 行（amount_mop + 調整前/調整/調整後 全 0）。對任何 sum 貢獻 0 → tie 不變（user 要 ensure tie）──
    _amtc = [c for c in ["amount_mop", "調整前_萬", "調整_萬", "調整後_萬"] if c in combined.columns]
    if _amtc and "entity" in combined.columns:
        _pre = combined.groupby("entity")["amount_mop"].apply(lambda s: pd.to_numeric(s, errors="coerce").sum()).to_dict()
        _az = pd.Series(True, index=combined.index)
        for c in _amtc:
            _az &= pd.to_numeric(combined[c], errors="coerce").fillna(0.0).eq(0.0)
        if int(_az.sum()):
            combined = combined[~_az].copy()
            _post = combined.groupby("entity")["amount_mop"].apply(lambda s: pd.to_numeric(s, errors="coerce").sum()).to_dict()
            print(f"  [剔全0行] drop {int(_az.sum()):,} 行（全部金額=0）；驗 tie：")
            for e in sorted(_pre):
                _d = _pre[e] - _post.get(e, 0.0)
                print(f"      {e:<7} Σ前={_pre[e]/1e6:>9,.1f}M  Σ後={_post.get(e,0)/1e6:>9,.1f}M  Δ={_d/1e6:.4f}M")
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
