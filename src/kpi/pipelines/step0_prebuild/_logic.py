"""Step 0 Prebuild: Multi-year extract + unify heterogeneous source files.

Supports `prebuild_years` config block in parameters.yml. Each year defines its own
sources + sheets, but all extracted rows share the same UNIFIED_COLS schema PLUS a
`report_year` column.

YAML structure:
  prebuild_years:
    - year: 2025
      raw_subdir: ""                    # files at data/<alias>/raw/
      sources: {f1_capex: ..., f3_gaming: ..., f4_nongaming: ..., f2_plan: ...}
      sheet_overrides: {f1_capex: "2025 Raw Data", ...}
      f1_budget_source_keep: "..."
    - year: 2024
      raw_subdir: ""
      sources: {je_main: "2024年JE.xlsx", retendering_pay: "2024/..."}
      sheet_extract: {capex: ["Capex"], gaming_opex: ["2024 Journal line (For KPMG)"], ...}
    - year: 2023
      raw_subdir: "2023"
      sources: {capex: "CAPEX.xlsx", gaming: "MGM-gaming.xlsx",
                opex: "OPEX.xlsx", payroll: "payroll.xlsx"}
      sheet_extract: {capex: ["JL details"], gaming_opex: ["opex 明細"], ...}

Backward compat: if `prebuild_years` not set, falls back to legacy `prebuild_sources`
(treated as single-year 2025).
"""
from __future__ import annotations

import logging
import re
from pathlib import Path

import pandas as pd

from kpi.lib.conf import load_config, path
from kpi.lib.io_setup import force_unbuffered_io

force_unbuffered_io()
log = logging.getLogger(__name__)

UNIFIED_COLS = [
    "source", "wd", "company", "category",
    "project_id", "project_name",
    "section", "nature", "item_type", "sub_type",
    "period", "amount_mop",
    "journal", "ledger_l4", "ledger_account",
    "journal_source", "supplier", "line_memo",
    "report_year",
]


# ── Helpers ───────────────────────────────────────────────────────────────────

def _load_xl(raw_dir: Path, fname: str, sheet, header: int = 0) -> pd.DataFrame:
    fpath = raw_dir / fname
    if not fpath.exists():
        raise FileNotFoundError(
            f"Prebuild source not found: {fpath}\n"
            f"Place the file in {raw_dir} before running this pipeline step."
        )
    log.info("  Reading %s::%s ...", fname, sheet)
    df = pd.read_excel(fpath, sheet_name=sheet, header=header, dtype=str, engine="openpyxl")
    log.info("    %s rows, %s cols", f"{len(df):,}", len(df.columns))
    return df


def _to_mop(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str).str.replace(",", "").str.strip(), errors="coerce"
    )


def _extract_project_id(s) -> str | None:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return None
    m = re.search(r'項目0*(\d+)', str(s))
    return m.group(1) if m else None


def _first_col(row: pd.Series, candidates: list[str], default: str = "") -> str:
    """First non-empty value from candidate column names."""
    for c in candidates:
        if c in row.index:
            v = row[c]
            if v is not None and not (isinstance(v, float) and pd.isna(v)) and str(v).strip() != "":
                return v
    return default


def _empty() -> pd.DataFrame:
    return pd.DataFrame(columns=UNIFIED_COLS)


# ── Generic JE extractor (works across years) ─────────────────────────────────

AMT_COL_CANDIDATES = ["Debit minus Credit", "Amount", "Transaction Debit minus Credit"]
PROJ_ID_COL_CANDIDATES = ["項目", "項目序號1", "承批公司項目序號"]
PROJ_NAME_COL_CANDIDATES = ["Item Name", "項目名稱", "Event Name", "Type"]
SECTION_COL_CANDIDATES = ["Section", "Main Group", "Section No."]
NATURE_COL_CANDIDATES = ["Nature", "Sub Group - Chinese", "Section Nature"]
SUBTYPE_COL_CANDIDATES = ["Ledger Hierarchy Level 5", "Spend Category as Worktag", "Type"]
LINE_MEMO_CANDIDATES = ["Journal Line Memo", "Line Item Description", "Memo"]


def _extract_je_generic(
    raw_dir: Path, fname: str, sheet: str,
    source_tag: str, wd: str, item_type: str, year: int,
    *,
    category: str | list[str] | None = None,
    apply_offset_filter: bool = False,
    apply_reversed_filter: bool = False,
    apply_construction_in_progress_filter: bool = False,
    budget_source_keep: str | None = None,
) -> pd.DataFrame:
    """Row-level JE extractor with flexible column lookup + opt-in filters."""
    log.info("[%s y=%s] %s::%s", source_tag, year, fname, sheet)
    try:
        df = _load_xl(raw_dir, fname, sheet)
    except FileNotFoundError as e:
        log.warning("  %s", e)
        return _empty()
    n0 = len(df)
    if n0 == 0:
        return _empty()

    # Filters (case-insensitive where applicable)
    if category and "Category" in df.columns:
        cats = [category] if isinstance(category, str) else list(category)
        df = df[df["Category"].fillna("").astype(str).isin(cats)]
    if apply_offset_filter:
        offset_cols = []
        for c in df.columns:
            cl = str(c).strip().lower()
            if cl == "offset" or ("filter" in cl and "offset" in cl):
                offset_cols.append(c)
        if not offset_cols and "Remarks" in df.columns:
            offset_cols = ["Remarks"]
        for c in offset_cols:
            df = df[~df[c].fillna("").astype(str).str.strip().str.lower().isin({"y", "offset"})]
    if apply_construction_in_progress_filter and "Ledger Hierarchy Level 4" in df.columns:
        df = df[~df["Ledger Hierarchy Level 4"].fillna("").astype(str).str.contains(
            "Construction In Progress", na=False
        )]
    if apply_reversed_filter:
        for c in ("Reversed", "Reversed?"):
            if c in df.columns:
                df = df[df[c].fillna("").astype(str).str.strip().str.lower() != "yes"]
                break
    if budget_source_keep and "Budget Source" in df.columns:
        before = len(df)
        df = df[df["Budget Source"].fillna("").astype(str).str.strip() == budget_source_keep]
        log.info("  Budget Source filter (%s): %s → %s", budget_source_keep, before, len(df))

    log.info("  Filtered: %s → %s rows", n0, len(df))
    if len(df) == 0:
        return _empty()

    rows = []
    for _, r in df.iterrows():
        proj_id_src = _first_col(r, PROJ_ID_COL_CANDIDATES)
        proj_name = str(_first_col(r, PROJ_NAME_COL_CANDIDATES) or "")
        cat_val = r.get("Category", "") if "Category" in r.index else ""
        if (not cat_val) and isinstance(category, str):
            cat_val = category
        rows.append({
            "source": f"{source_tag}_{year}",
            "wd": wd,
            "company": r.get("Company", "") if "Company" in r.index else "",
            "category": str(cat_val or ""),
            "project_id": _extract_project_id(proj_id_src),
            "project_name": proj_name,
            "section": str(_first_col(r, SECTION_COL_CANDIDATES) or ""),
            "nature": str(_first_col(r, NATURE_COL_CANDIDATES) or ""),
            "item_type": item_type,
            "sub_type": str(_first_col(r, SUBTYPE_COL_CANDIDATES) or ""),
            "period": r.get("Period", "") if "Period" in r.index else "",
            "amount_mop": _first_col(r, AMT_COL_CANDIDATES),
            "journal": r.get("Journal", "") if "Journal" in r.index else "",
            "ledger_l4": r.get("Ledger Hierarchy Level 4", "") if "Ledger Hierarchy Level 4" in r.index else "",
            "ledger_account": r.get("Ledger Account", "") if "Ledger Account" in r.index else "",
            "journal_source": r.get("Journal Source", "") if "Journal Source" in r.index else "",
            "supplier": (r.get("Supplier", "") if "Supplier" in r.index else "")
                         or (r.get("Vendor", "") if "Vendor" in r.index else ""),
            "line_memo": str(_first_col(r, LINE_MEMO_CANDIDATES) or ""),
            "report_year": year,
        })
    out = pd.DataFrame(rows, columns=UNIFIED_COLS)
    out["amount_mop"] = _to_mop(out["amount_mop"])
    log.info("  %s rows  MOP total: %s", len(out), f"{out['amount_mop'].sum():,.0f}")
    return out


# ── Pivot extractors (Payroll / COGS pivots that need melting) ───────────────

def _extract_payroll_pivot(
    raw_dir: Path, fname: str, sheet: str,
    source_tag: str, wd: str, year: int,
) -> pd.DataFrame:
    log.info("[%s y=%s] payroll pivot %s::%s", source_tag, year, fname, sheet)
    try:
        df = _load_xl(raw_dir, fname, sheet)
    except FileNotFoundError as e:
        log.warning("  %s", e)
        return _empty()
    month_cols = [c for c in df.columns if re.match(r"^\d{4}\s*-\s*\w{3,}$", str(c))]
    if not month_cols:
        log.warning("  No month columns found")
        return _empty()
    rl = "Row Labels" if "Row Labels" in df.columns else df.columns[0]
    df = df[[rl] + month_cols].copy()
    df[rl] = df[rl].fillna("").astype(str)
    df = df[df[rl].str.match(r"^\d{6}")]
    melted = df.melt(id_vars=rl, value_vars=month_cols, var_name="period", value_name="amount_mop")
    melted["amount_mop"] = _to_mop(melted["amount_mop"])
    melted = melted[melted["amount_mop"].notna() & (melted["amount_mop"] != 0)]
    rows = []
    for _, r in melted.iterrows():
        dept = re.sub(r"^\d{6}\s*", "", str(r[rl])).strip()
        rows.append({
            "source": f"{source_tag}_{year}",
            "wd": wd, "company": "", "category": "Non-Gaming",
            "project_id": None, "project_name": f"人工成本 – {dept}",
            "section": "", "nature": "人工成本",
            "item_type": "Payroll", "sub_type": dept,
            "period": r["period"], "amount_mop": r["amount_mop"],
            "journal": "", "ledger_l4": "Standard: Payroll(L4)",
            "ledger_account": str(r[rl]),
            "journal_source": "", "supplier": "", "line_memo": dept,
            "report_year": year,
        })
    out = pd.DataFrame(rows, columns=UNIFIED_COLS)
    log.info("  %s rows  MOP total: %s", len(out), f"{out['amount_mop'].sum():,.0f}")
    return out


def _extract_cogs_pivot(
    raw_dir: Path, fname: str, sheet: str,
    source_tag: str, wd: str, year: int,
) -> pd.DataFrame:
    log.info("[%s y=%s] cogs pivot %s::%s", source_tag, year, fname, sheet)
    try:
        df = _load_xl(raw_dir, fname, sheet)
    except FileNotFoundError as e:
        log.warning("  %s", e)
        return _empty()
    month_cols = [c for c in df.columns
                  if re.match(r"^\d{4}\s*-\s*\w{3,}$", str(c)) and not str(c).endswith(".1")]
    if not month_cols:
        log.warning("  No month columns found")
        return _empty()
    rl = "Row Labels" if "Row Labels" in df.columns else df.columns[0]
    df = df[[rl] + month_cols].copy()
    df[rl] = df[rl].fillna("").astype(str)
    df = df[df[rl].str.match(r"^\d{6}")]
    melted = df.melt(id_vars=rl, value_vars=month_cols, var_name="period", value_name="amount_mop")
    melted["amount_mop"] = _to_mop(melted["amount_mop"])
    melted = melted[melted["amount_mop"].notna() & (melted["amount_mop"] > 0)]
    rows = []
    for _, r in melted.iterrows():
        outlet = re.sub(r"^\d{6}\s*", "", str(r[rl])).strip()
        rows.append({
            "source": f"{source_tag}_{year}",
            "wd": wd, "company": "", "category": "Non-Gaming",
            "project_id": None, "project_name": f"餐飲收支 – {outlet}",
            "section": "", "nature": "美食之都",
            "item_type": "COGS", "sub_type": outlet,
            "period": r["period"], "amount_mop": r["amount_mop"],
            "journal": "", "ledger_l4": "Standard: Cost of Sales(L4)",
            "ledger_account": str(r[rl]),
            "journal_source": "", "supplier": "", "line_memo": outlet,
            "report_year": year,
        })
    out = pd.DataFrame(rows, columns=UNIFIED_COLS)
    log.info("  %s rows  MOP total: %s", len(out), f"{out['amount_mop'].sum():,.0f}")
    return out


def _extract_patron(
    raw_dir: Path, fname: str, sheet: str,
    source_tag: str, year: int,
) -> pd.DataFrame:
    log.info("[%s y=%s] patron %s::%s", source_tag, year, fname, sheet)
    try:
        df = _load_xl(raw_dir, fname, sheet)
    except FileNotFoundError as e:
        log.warning("  %s", e)
        return _empty()
    if "Amount" in df.columns:
        df["Amount"] = _to_mop(df["Amount"])
        df = df[df["Amount"].notna() & (df["Amount"] != 0)]
    elif "MOP" in df.columns:
        df["Amount"] = _to_mop(df["MOP"])
        df = df[df["Amount"].notna() & (df["Amount"] != 0)]
    rows = []
    for _, r in df.iterrows():
        event = str(r.get("Event Name", "") if "Event Name" in r.index else
                    r.get("項目名稱", "") if "項目名稱" in r.index else "")
        rows.append({
            "source": f"{source_tag}_{year}",
            "wd": "WD5_Patron", "company": "", "category": "Non-Gaming",
            "project_id": _extract_project_id(
                event or (r.get("承批公司項目序號", "") if "承批公司項目序號" in r.index else "")
            ),
            "project_name": event,
            "section": str(r.get("Section", "") if "Section" in r.index else ""),
            "nature": str(r.get("Nature", "") if "Nature" in r.index else ""),
            "item_type": "OPEX",
            "sub_type": str(r.get("Type", "") if "Type" in r.index else
                            r.get("Item Type", "") if "Item Type" in r.index else ""),
            "period": str(r.get("Accounting Date", "") if "Accounting Date" in r.index else
                          r.get("Period", "") if "Period" in r.index else ""),
            "amount_mop": r.get("Amount", "") if "Amount" in r.index else "",
            "journal": str(r.get("Comp ID", "") if "Comp ID" in r.index else ""),
            "ledger_l4": "",
            "ledger_account": str(r.get("Item Type", "") if "Item Type" in r.index else ""),
            "journal_source": str(r.get("Promotion Type", "") if "Promotion Type" in r.index else ""),
            "supplier": "",
            "line_memo": str(r.get("Promotion", "") if "Promotion" in r.index else ""),
            "report_year": year,
        })
    out = pd.DataFrame(rows, columns=UNIFIED_COLS)
    out["amount_mop"] = _to_mop(out["amount_mop"])
    log.info("  %s rows  MOP total: %s", len(out), f"{out['amount_mop'].sum():,.0f}")
    return out


# ── Plan summary (reference only) ─────────────────────────────────────────────

def _extract_f2_plan(raw_dir: Path, fname: str, sheet: str = "Master") -> pd.DataFrame:
    log.info("[plan] %s", fname)
    try:
        df = _load_xl(raw_dir, fname, sheet, header=5)
    except FileNotFoundError as e:
        log.warning("  %s", e)
        return pd.DataFrame()
    df = df[[c for c in df.columns if not str(c).startswith("Unnamed")]]
    rename = {
        "項目序號": "project_id", "項目名稱:": "project_name",
        "Payroll": "actual_payroll", "CAPEX": "actual_capex",
        "OPEX": "actual_opex", "Total": "actual_total",
        "2025年度計劃": "plan_2025", "Remarks": "remarks",
    }
    df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})
    if "project_id" in df.columns:
        df = df[df["project_id"].notna() & (df["project_id"].astype(str).str.strip() != "")]
    log.info("  %s projects in plan", len(df))
    return df


# ── Year orchestrators ────────────────────────────────────────────────────────

def _prebuild_2025(raw_dir: Path, sources: dict, sheet_overrides: dict,
                   budget_source_keep: str | None, keep_cip: bool = False) -> pd.DataFrame:
    parts = []
    f1 = sources.get("f1_capex") or sources.get("f1") or ""
    f3 = sources.get("f3_gaming") or sources.get("f3") or ""
    f4 = sources.get("f4_nongaming") or sources.get("f4") or ""
    if f1:
        sheet = sheet_overrides.get("f1_capex") or sheet_overrides.get("f1") or "2025 Raw Data"
        # keep_cip=True keeps Construction-In-Progress rows (the construction CAPEX the
        # project-team golden counts as investment); default drops them (legacy behaviour).
        parts.append(_extract_je_generic(
            raw_dir, f1, sheet, "S1_CAPEX", "CAPEX", "CAPEX", 2025,
            category=["Gaming", "Non-Gaming"],
            apply_offset_filter=True,
            apply_construction_in_progress_filter=(not keep_cip),
            apply_reversed_filter=True,
            budget_source_keep=budget_source_keep,
        ))
    if f3:
        sheet = sheet_overrides.get("f3_gaming") or sheet_overrides.get("f3") or "Data - JEs"
        parts.append(_extract_je_generic(
            raw_dir, f3, sheet, "S3_GamingOpex", "Gaming_OPEX", "OPEX", 2025,
            category="Gaming",
            apply_offset_filter=True,
        ))
    if f4:
        parts.append(_extract_je_generic(
            raw_dir, f4, "WD1", "S4_WD1", "WD1", "OPEX", 2025,
            category="Non-Gaming", apply_offset_filter=True,
        ))
        parts.append(_extract_je_generic(
            raw_dir, f4, "WD2-Allocation", "S4_WD2", "WD2", "OPEX", 2025,
            category="Non-Gaming",
        ))
        parts.append(_extract_payroll_pivot(raw_dir, f4, "WD3-Payroll", "S4_WD3", "WD3", 2025))
        parts.append(_extract_cogs_pivot(raw_dir, f4, "WD4-COGS", "S4_WD4", "WD4", 2025))
        parts.append(_extract_patron(raw_dir, f4, "Patron Management", "S4_Patron", 2025))
    filtered = [p for p in parts if len(p) > 0]
    return pd.concat(filtered, ignore_index=True) if filtered else _empty()


def _prebuild_2024(raw_dir: Path, sources: dict, sheet_extract: dict) -> pd.DataFrame:
    parts = []
    je_main = sources.get("je_main", "")
    if not je_main:
        log.warning("  2024: no je_main configured")
        return _empty()
    for sheet in sheet_extract.get("capex") or ["Capex"]:
        parts.append(_extract_je_generic(
            raw_dir, je_main, sheet, "S_Capex", "CAPEX", "CAPEX", 2024,
            apply_construction_in_progress_filter=True,
            apply_reversed_filter=True,
        ))
    for sheet in sheet_extract.get("gaming_opex") or ["2024 Journal line (For KPMG)"]:
        parts.append(_extract_je_generic(
            raw_dir, je_main, sheet, "S_GamingOpex", "Gaming_OPEX", "OPEX", 2024,
            apply_reversed_filter=True,
        ))
    for sheet in sheet_extract.get("nongaming_wd1") or ["WD1"]:
        parts.append(_extract_je_generic(
            raw_dir, je_main, sheet, "S_WD1", "WD1", "OPEX", 2024,
            category="Non-Gaming", apply_offset_filter=True,
        ))
    for sheet in sheet_extract.get("nongaming_wd2") or ["WD2-Allocation"]:
        parts.append(_extract_je_generic(
            raw_dir, je_main, sheet, "S_WD2", "WD2", "OPEX", 2024,
            category="Non-Gaming",
        ))
    for sheet in sheet_extract.get("nongaming_payroll_pivot") or ["WD3-Payroll"]:
        parts.append(_extract_payroll_pivot(raw_dir, je_main, sheet, "S_WD3", "WD3", 2024))
    for sheet in sheet_extract.get("nongaming_cogs_pivot") or ["WD4-COGS"]:
        parts.append(_extract_cogs_pivot(raw_dir, je_main, sheet, "S_WD4", "WD4", 2024))
    for sheet in sheet_extract.get("patron_mgmt") or ["PM"]:
        parts.append(_extract_patron(raw_dir, je_main, sheet, "S_Patron", 2024))
    filtered = [p for p in parts if len(p) > 0]
    return pd.concat(filtered, ignore_index=True) if filtered else _empty()


def _prebuild_2023(raw_dir: Path, sources: dict, sheet_extract: dict) -> pd.DataFrame:
    parts = []
    capex_f = sources.get("capex", "")
    gaming_f = sources.get("gaming", "")
    opex_f = sources.get("opex", "")
    if capex_f:
        for sheet in sheet_extract.get("capex") or ["JL details"]:
            parts.append(_extract_je_generic(
                raw_dir, capex_f, sheet, "S_Capex", "CAPEX", "CAPEX", 2023,
                apply_reversed_filter=True,
            ))
    if gaming_f:
        for sheet in sheet_extract.get("gaming_opex") or ["opex 明細"]:
            parts.append(_extract_je_generic(
                raw_dir, gaming_f, sheet, "S_GamingOpex", "Gaming_OPEX", "OPEX", 2023,
                category="Gaming",
            ))
    if opex_f:
        for sheet in sheet_extract.get("nongaming_wd1") or ["WD#1"]:
            parts.append(_extract_je_generic(
                raw_dir, opex_f, sheet, "S_WD1", "WD1", "OPEX", 2023,
                category="Non-Gaming", apply_offset_filter=True,
            ))
        for sheet in sheet_extract.get("nongaming_wd2") or ["WD#2"]:
            parts.append(_extract_je_generic(
                raw_dir, opex_f, sheet, "S_WD2", "WD2", "OPEX", 2023,
                category="Non-Gaming",
            ))
        # 2023 payroll is JE-row (Data#5), NOT pivot — use generic JE extractor
        for sheet in sheet_extract.get("nongaming_payroll_je") or ["Data#5"]:
            parts.append(_extract_je_generic(
                raw_dir, opex_f, sheet, "S_Payroll", "WD3", "Payroll", 2023,
                apply_reversed_filter=True,
            ))
        # 2023 cogs is also JE-row (Data#6)
        for sheet in sheet_extract.get("nongaming_cogs_je") or ["Data#6"]:
            parts.append(_extract_je_generic(
                raw_dir, opex_f, sheet, "S_COGS", "WD4", "COGS", 2023,
            ))
        for sheet in sheet_extract.get("patron_mgmt") or ["PM"]:
            parts.append(_extract_patron(raw_dir, opex_f, sheet, "S_Patron", 2023))
    filtered = [p for p in parts if len(p) > 0]
    return pd.concat(filtered, ignore_index=True) if filtered else _empty()


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    cfg = load_config()
    company = cfg["company"]["code"]
    master = cfg["_master"]
    ent = (master.get("companies") or {}).get(company, {})

    raw_dir = path(cfg, "interim_dir").parent / "raw"
    raw_dir.mkdir(parents=True, exist_ok=True)

    # Multi-year config OR fallback to legacy single-year (2025)
    years_cfg = ent.get("prebuild_years")
    if not years_cfg:
        legacy_sources = ent.get("prebuild_sources") or {}
        if not legacy_sources:
            raise ValueError(
                f"No prebuild_years or prebuild_sources configured for {company}. "
                f"Add `prebuild_years` (preferred) or `prebuild_sources` in "
                f"conf/{company}/parameters.yml."
            )
        years_cfg = [{
            "year": 2025,
            "raw_subdir": "",
            "sources": {
                "f1_capex": legacy_sources.get("f1", ""),
                "f3_gaming": legacy_sources.get("f3", ""),
                "f4_nongaming": legacy_sources.get("f4", ""),
                "f2_plan": legacy_sources.get("f2", ""),
            },
            "sheet_overrides": {
                "f1_capex": ent.get("prebuild_sheet_f1", "2025 Raw Data"),
                "f3_gaming": ent.get("prebuild_sheet_f3", "Data - JEs"),
                "f2_plan": ent.get("prebuild_sheet_f2", "Master"),
            },
            "f1_budget_source_keep": ent.get("prebuild_f1_budget_source_keep"),
        }]

    log.info("Prebuild: %s year(s) configured", len(years_cfg))

    yearly_parts: list[pd.DataFrame] = []
    plan_args = None  # (raw_dir, fname, sheet) — taken from first year with f2_plan

    for ycfg in years_cfg:
        year = ycfg["year"]
        subdir = ycfg.get("raw_subdir", "") or ""
        year_raw_dir = (raw_dir / subdir) if subdir else raw_dir
        sources = ycfg.get("sources", {})
        sheet_overrides = ycfg.get("sheet_overrides", {})
        sheet_extract = ycfg.get("sheet_extract", {})

        log.info("=" * 70)
        log.info("Year %s (raw=%s)", year, year_raw_dir)
        log.info("=" * 70)

        if year == 2025:
            df_y = _prebuild_2025(year_raw_dir, sources, sheet_overrides,
                                  ycfg.get("f1_budget_source_keep"),
                                  keep_cip=ent.get("prebuild_keep_cip", False))
            f2 = sources.get("f2_plan") or sources.get("f2")
            if f2 and plan_args is None:
                plan_sheet = (sheet_overrides.get("f2_plan")
                              or sheet_overrides.get("f2")
                              or "Master")
                plan_args = (year_raw_dir, f2, plan_sheet)
        elif year == 2024:
            df_y = _prebuild_2024(year_raw_dir, sources, sheet_extract)
        elif year == 2023:
            df_y = _prebuild_2023(year_raw_dir, sources, sheet_extract)
        else:
            log.warning("Year %s not supported (only 2023/2024/2025)", year)
            continue

        log.info("Year %s subtotal: %s rows  MOP %s",
                 year, f"{len(df_y):,}",
                 f"{df_y['amount_mop'].sum():,.0f}" if len(df_y) else "0")
        yearly_parts.append(df_y)

    actuals = pd.concat([y for y in yearly_parts if len(y) > 0], ignore_index=True)
    log.info("=" * 70)
    log.info("Combined: %s rows", f"{len(actuals):,}")
    if len(actuals) > 0:
        log.info("By year:\n%s",
                 actuals.groupby("report_year")["amount_mop"]
                 .agg(count="count", total="sum").round(0))
        log.info("By source:\n%s",
                 actuals.groupby("source")["amount_mop"]
                 .agg(count="count", total="sum").round(0))

    # ledger_l4 exclude filter (applied across all years)
    exclude_l4 = ent.get("prebuild_exclude_ledger_l4") or []
    if exclude_l4 and "ledger_l4" in actuals.columns:
        before = len(actuals)
        mask = actuals["ledger_l4"].fillna("").isin(exclude_l4)
        actuals = actuals[~mask].reset_index(drop=True)
        log.info("Excluded %s rows by ledger_l4 filter: %s → %s",
                 before - len(actuals), before, len(actuals))

    raw_filename = cfg["paths"]["raw_filename"]
    out_path = raw_dir / raw_filename
    actuals.to_parquet(out_path, index=False)
    log.info("Wrote actuals → %s  (%.1f MB, %s rows)",
             out_path.name, out_path.stat().st_size / 1e6, f"{len(actuals):,}")

    if plan_args:
        plan_dir, plan_fname, plan_sheet = plan_args
        plan = _extract_f2_plan(plan_dir, plan_fname, plan_sheet)
        if len(plan) > 0:
            stem = Path(raw_filename).stem
            plan_path = raw_dir / f"{stem}_plan.parquet"
            plan.to_parquet(plan_path, index=False)
            log.info("Wrote plan → %s", plan_path.name)


if __name__ == "__main__":
    main()
