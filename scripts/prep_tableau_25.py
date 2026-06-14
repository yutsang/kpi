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

def run(fmt="per-entity-xlsx", out_dir="data/tableau"):
    """Build Tableau files. per-entity-xlsx → out_dir/tableau_{yr}_{ent}.xlsx (default data/tableau).
    Importable so the `generate` kedro pipeline can call it (no argparse)."""
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
        # NG0-NG11 + Chinese label derived from vertical_id via canonical V→NG mapping.
        # This gives consistent NG codes across all 6 entities (raw data for non-Galaxy
        # entities only has binary gaming/non_gaming — we derive granular NG from V).
        if "vertical_id" in df.columns:
            df["ng_code"] = df["vertical_id"].map(lambda v: V_TO_NG.get(str(v), ("NG11", "其他"))[0])
            df["ng_label"] = df["vertical_id"].map(lambda v: V_TO_NG.get(str(v), ("NG11", "其他"))[1])
            keep.append("ng_code")
            keep.append("ng_label")
        # Per-entity native cols (preserve project + subproject + acct for drill-down)
        proj_col = cfg.get("columns",{}).get("project","")
        ac_col = cfg.get("columns",{}).get("account_code","")
        ad_col = cfg.get("columns",{}).get("account_desc","")
        dn_col = cfg.get("columns",{}).get("description","")
        vd_col = cfg.get("columns",{}).get("vendor","")
        # subproject candidates (varies per entity)
        sub_col = next((c for c in ("Sub project", "SubProject_Name", "Subproject_Name",
                                      "subproject", "项目名称中文", "项目英文名称",
                                      "Initiative Name", "Contents Name") if c in df.columns), None)
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
                   "take_flag", "take_flag2", "netoff_flag", "internal", "remark"):
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
    print(f"Cols: {list(combined.columns)}")

    # ── cube / cube-detail: ONE aggregated file = cross-tab source (no union, no stitching) ──
    if fmt in ("cube", "cube-detail"):
        dims = ["entity", "year_bucket", "ng_code", "ng_label",
                "vertical_id", "vertical_label", "horizontal_id", "horizontal_label",
                "ng_scope", "final_capex_opex"]
        if fmt == "cube-detail":   # keep drill-down dims (project / account / vendor)
            dims += ["project", "subproject", "account_code", "account_desc", "vendor"]
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
                   default="per-entity-xlsx",
                   help="cube/cube-detail = aggregated cross-tab; csv = ONE row-level CSV (all dims incl description); "
                        "csv-per-entity = 6 row-level CSVs (one per company, for 對數); per-entity-xlsx (default) = old union")
    p.add_argument("--out", default="data/tableau", help="output dir for per-entity-xlsx")
    args = p.parse_args()
    run(args.format, args.out)


if __name__ == "__main__":
    main()
