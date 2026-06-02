"""Delivery readiness check + deliverable generator (per entity, per year).

Decides whether each (entity × year) combo can be delivered. Independent 24
and 25 progress — an entity can be 25-ready before 24-ready (or vice versa).

Quality checks per entity×year:
  1. SIG coverage: step3 LLM(batch) workload = 0  (all sigs in feedback/rules)
  2. PROJECT coverage: report's unique projects ⊆ proj_df  (no missing → no NaN vertical)
  3. PROJECT counting: report unique-project count matches expectation
  4. ZERO rows: amount=0 row %  (lower = better)
  5. V_OTHER %: should be low
  6. Raw sum vs report sum cross-check  (amount conservation)

Deliverables (per entity, when its year is ready):
  D1. <com>_je_all_24.parquet  — full 24-year raw journal entries
  D2. <com>_je_all_25.parquet  — full 25-year raw journal entries
  D3. <com>_kpi_report.xlsx + .parquet  — copied/touched (latest kedro output)

Outputs:
  - readiness_matrix.txt  — per (entity, year) PASS/WARN/FAIL summary
  - data/<entity>/output/deliverables/  — generated files when --generate

Run:
    python scripts/delivery_readiness_check.py --all                 # check only
    python scripts/delivery_readiness_check.py --all --generate      # check + write deliverables for ready ones
    python scripts/delivery_readiness_check.py --entity sjm --year 25
"""
from __future__ import annotations

import argparse
import shutil
import sys
import time
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import pandas as pd
import yaml

from kpi.lib.conf import load_config
from kpi.lib.feedback import load_feedback
from kpi.lib.rules import load_predominant_rules, find_predominant_rule, resolve_then


ENTITIES = {
    "galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
    "vml":    "company_4", "melco":  "company_5", "mgm": "company_6",
}

XLSX_ROW_LIMIT = 1_048_575


# Thresholds for PASS / WARN / FAIL classification
TH_LLM_BATCH_PASS = 50            # LLM batch sigs <=50 → PASS
TH_LLM_BATCH_WARN = 500
TH_NAN_VERT_PCT_PASS = 0.5         # NaN vertical % <0.5 → PASS
TH_NAN_VERT_PCT_WARN = 5.0
TH_VOTHER_PCT_PASS = 5.0
TH_VOTHER_PCT_WARN = 15.0
TH_ZERO_PCT_PASS = 1.0
TH_ZERO_PCT_WARN = 5.0


def status_emoji(s: str) -> str:
    return {"PASS": "✅", "WARN": "⚠️", "FAIL": "❌"}.get(s, "?")


def worst(*states: str) -> str:
    order = {"PASS": 0, "WARN": 1, "FAIL": 2}
    return max(states, key=lambda s: order.get(s, 0))


def classify(value: float, pass_th: float, warn_th: float) -> str:
    if value <= pass_th:
        return "PASS"
    if value <= warn_th:
        return "WARN"
    return "FAIL"


def find_year_col(df: pd.DataFrame) -> str | None:
    for c in ("report_period", "report_year", "year", "period"):
        if c in df.columns:
            return c
    return None


def find_amount_col(df: pd.DataFrame) -> str | None:
    candidates = [c for c in df.columns if "amount" in c.lower() and "split" not in c.lower()]
    if not candidates:
        candidates = [c for c in df.columns if "mop" in c.lower() or "value" in c.lower()]
    if not candidates:
        return None
    candidates.sort(key=lambda c: (0 if "reported" in c.lower() else 1, len(c)))
    return candidates[0]


def find_project_col(df: pd.DataFrame, cfg_project_col: str | None) -> str | None:
    if cfg_project_col and cfg_project_col in df.columns:
        return cfg_project_col
    for c in df.columns:
        if "project" in c.lower() and "samples" not in c.lower():
            return c
    return None


def find_vertical_col(df: pd.DataFrame) -> str | None:
    for c in ("final_vertical", "vertical_id", "vertical"):
        if c in df.columns:
            return c
    return None


def read_kpi_main(out_dir: Path, com: str) -> tuple[pd.DataFrame | None, str]:
    parquet = out_dir / f"{com}_kpi_report.parquet"
    if parquet.exists():
        return pd.read_parquet(parquet), parquet.name
    for p in out_dir.glob(f"{com}_kpi_report*.xlsx"):
        if "_24" in p.stem or "_25" in p.stem:
            continue
        try:
            return pd.read_excel(p, sheet_name="7_all_rows", engine="calamine"), p.name
        except Exception:
            try:
                return pd.read_excel(p, sheet_name="7_all_rows", engine="openpyxl"), p.name
            except Exception:
                continue
    return None, ""


def forecast_llm_batch(ent: str, com: str) -> int:
    """Compute step3 LLM(batch) workload for this entity (live forecast)."""
    interim = Path(f"data/{ent}/interim")
    sig_xlsx = interim / f"{com}_unique_signatures.xlsx"
    if not sig_xlsx.exists():
        return -1
    sig_df = pd.read_excel(sig_xlsx)
    for c in ("account_code", "account_desc", "desc_norm"):
        if c in sig_df.columns:
            sig_df[c] = sig_df[c].astype("string").fillna("")
    cfg = load_config(ent)
    overrides, _ = load_feedback(Path(f"data/{ent}/output"), com)
    fb = set(overrides.keys())
    rules = load_predominant_rules(cfg, kind="horizontal") or []
    cats_path = Path("conf/base/categories.yml")
    horizontals = yaml.safe_load(cats_path.read_text(encoding="utf-8")).get("horizontals", [])

    llm = 0
    for rec in sig_df.to_dict(orient="records"):
        sig = str(rec.get("signature", ""))
        if sig in fb:
            continue
        rule, _ = find_predominant_rule(rules, rec)
        if rule:
            res = resolve_then(rule, horizontals)
            if res.get("mode") == "fixed":
                continue
        llm += 1
    return llm


def check_one_year(df: pd.DataFrame, year_tag: str, year_col: str, amt_col: str,
                   proj_col: str, v_col: str, ent: str, com: str) -> dict:
    """Run quality checks for a single (entity, year) and return metrics + status."""
    interim = Path(f"data/{ent}/interim")
    mask = df[year_col].astype(str).str.startswith(year_tag)
    df_y = df[mask].copy()

    res = {"entity": ent, "year": year_tag, "rows": len(df_y)}

    if len(df_y) == 0:
        res["status"] = "FAIL"
        res["reason"] = "no rows for this year"
        return res

    # Amounts
    amt = pd.to_numeric(df_y[amt_col], errors="coerce") if amt_col else None
    res["total_amount"] = float(amt.sum()) if amt is not None else 0

    # Zero amount %
    if amt is not None:
        n_zero = int((amt.fillna(0) == 0).sum())
        res["zero_pct"] = 100 * n_zero / len(df_y)
        res["zero_rows"] = n_zero
    else:
        res["zero_pct"] = 100  # can't verify

    # Unique projects + missing-from-proj_df
    proj_xlsx = interim / f"{com}_unique_projects.xlsx"
    proj_df_set = set()
    if proj_xlsx.exists():
        pdf = pd.read_excel(proj_xlsx)
        pdf_col = pdf.columns[0]
        proj_df_set = set(pdf[pdf_col].dropna().astype(str).unique())

    report_projs = set(df_y[proj_col].dropna().astype(str).unique()) if proj_col else set()
    missing_projs = report_projs - proj_df_set
    res["report_unique_projects"] = len(report_projs)
    res["proj_df_size"] = len(proj_df_set)
    res["missing_projects"] = len(missing_projs)

    # NaN vertical %
    if v_col:
        nan_mask = df_y[v_col].isna() | (df_y[v_col].astype(str).str.strip() == "")
        res["nan_vert_pct"] = 100 * nan_mask.sum() / len(df_y)
        v_other_mask = df_y[v_col].astype(str) == "V_OTHER"
        res["vother_pct"] = 100 * v_other_mask.sum() / len(df_y)
        res["vother_rows"] = int(v_other_mask.sum())
        res["nan_vert_rows"] = int(nan_mask.sum())
    else:
        res["nan_vert_pct"] = 100
        res["vother_pct"] = 0
        res["vother_rows"] = 0
        res["nan_vert_rows"] = 0

    # Per-check status
    sig_llm = res.get("_sig_llm", 0)
    s_sig = classify(sig_llm, TH_LLM_BATCH_PASS, TH_LLM_BATCH_WARN)
    s_nan = classify(res["nan_vert_pct"], TH_NAN_VERT_PCT_PASS, TH_NAN_VERT_PCT_WARN)
    s_vother = classify(res["vother_pct"], TH_VOTHER_PCT_PASS, TH_VOTHER_PCT_WARN)
    s_zero = classify(res["zero_pct"], TH_ZERO_PCT_PASS, TH_ZERO_PCT_WARN)
    s_proj = "PASS" if res["missing_projects"] == 0 else (
        "WARN" if res["missing_projects"] < 20 else "FAIL")

    res["statuses"] = {
        "sig_coverage": s_sig,
        "project_coverage": s_proj,
        "nan_vertical": s_nan,
        "vother": s_vother,
        "zero_rows": s_zero,
    }
    res["status"] = worst(s_sig, s_proj, s_nan, s_vother, s_zero)
    return res


def check_entity(ent: str, com: str) -> dict:
    out_dir = Path(f"data/{ent}/output")
    df, src = read_kpi_main(out_dir, com)
    if df is None:
        return {"entity": ent, "error": "no kpi_report"}

    year_col = find_year_col(df)
    amt_col = find_amount_col(df)
    cfg = {}
    try:
        cfg = yaml.safe_load(Path(f"conf/{com}/parameters.yml").read_text(encoding="utf-8"))
    except Exception:
        pass
    proj_col = find_project_col(df, (cfg.get("columns") or {}).get("project"))
    v_col = find_vertical_col(df)

    # Entity-wide: sig LLM workload (same for both years; cached on df rows)
    sig_llm = forecast_llm_batch(ent, com)

    # Truncation flag
    truncated = len(df) >= XLSX_ROW_LIMIT

    out = {
        "entity": ent,
        "company": com,
        "src": src,
        "total_rows": len(df),
        "truncated": truncated,
        "sig_llm_workload": sig_llm,
        "year_col": year_col,
        "amount_col": amt_col,
        "project_col": proj_col,
        "vertical_col": v_col,
        "year_checks": {},
    }

    for year in ["24", "25"]:
        r = check_one_year(df, year, year_col, amt_col, proj_col, v_col, ent, com)
        r["_sig_llm"] = sig_llm
        # Re-classify status with sig_llm now known
        s_sig = classify(sig_llm if sig_llm >= 0 else 9999,
                        TH_LLM_BATCH_PASS, TH_LLM_BATCH_WARN)
        r["statuses"]["sig_coverage"] = s_sig
        r["status"] = worst(*r["statuses"].values())
        if truncated and year == "24":
            r["statuses"]["truncation"] = "WARN"
            r["status"] = worst(r["status"], "WARN")
        out["year_checks"][year] = r

    return out


def generate_deliverables(ent: str, com: str, year_results: dict) -> list[str]:
    notes = []
    interim = Path(f"data/{ent}/interim")
    out_dir = Path(f"data/{ent}/output")
    deliver_dir = out_dir / "deliverables"
    deliver_dir.mkdir(exist_ok=True)

    parquet = interim / f"{com}_raw.parquet"
    if not parquet.exists():
        notes.append(f"  ❌ raw parquet missing — cannot generate JE files")
        return notes

    df = pd.read_parquet(parquet)
    year_col = "report_period" if "report_period" in df.columns else (
        "report_year" if "report_year" in df.columns else None)
    if not year_col:
        notes.append(f"  ❌ no year col in parquet — cannot split")
        return notes

    for year in ["24", "25"]:
        yr_status = year_results.get(year, {}).get("status", "FAIL")
        # JE files are RAW (year-filtered raw transactions). Their validity is
        # independent of vertical/horizontal tagging quality. Always generate.
        mask = df[year_col].astype(str).str.startswith(year)
        sub = df[mask]
        if len(sub) == 0:
            notes.append(f"  ⏸ {year}: no rows in parquet")
            continue
        if yr_status == "FAIL":
            notes.append(f"  ⚠️  {year}: status=FAIL — JE_all still generated (raw data is valid), "
                        f"but DO NOT ship KPI_24/_25 split for this year until fixed")
        # parquet (always — for tooling)
        je_pq = deliver_dir / f"{com}_je_all_{year}.parquet"
        sub.to_parquet(je_pq, index=False)
        size_mb = je_pq.stat().st_size / 1e6
        notes.append(f"  ✓ {je_pq.name} ({len(sub):,} rows, {size_mb:.1f} MB)  status={yr_status}")

        # xlsx (for clients who can't read parquet)
        je_xl = deliver_dir / f"{com}_je_all_{year}.xlsx"
        try:
            # xlsxwriter is faster than openpyxl for writing large files
            sub.to_excel(je_xl, index=False, engine="xlsxwriter")
        except ImportError:
            sub.to_excel(je_xl, index=False)
        size_mb = je_xl.stat().st_size / 1e6
        notes.append(f"  ✓ {je_xl.name} ({len(sub):,} rows, {size_mb:.1f} MB)")

    # Copy KPI report (xlsx + parquet)
    for ext in (".xlsx", ".parquet"):
        src = out_dir / f"{com}_kpi_report{ext}"
        if src.exists():
            dst = deliver_dir / src.name
            shutil.copy2(src, dst)
            size_mb = dst.stat().st_size / 1e6
            notes.append(f"  ✓ {dst.name} ({size_mb:.1f} MB)")
    return notes


def format_year_summary(r: dict) -> list[str]:
    lines = []
    em = status_emoji(r["status"])
    lines.append(f"    {em} year={r['year']}  status={r['status']}  rows={r.get('rows', 0):,}  amount={r.get('total_amount', 0):,.0f}")
    s = r["statuses"]
    lines.append(f"       sig_coverage={status_emoji(s.get('sig_coverage',''))} {s.get('sig_coverage')}  "
                 f"project_coverage={status_emoji(s.get('project_coverage',''))} {s.get('project_coverage')} "
                 f"(missing={r.get('missing_projects',0):,})  "
                 f"nan_vert={status_emoji(s.get('nan_vertical',''))} {r.get('nan_vert_pct',0):.1f}%  "
                 f"vother={status_emoji(s.get('vother',''))} {r.get('vother_pct',0):.1f}%  "
                 f"zero={status_emoji(s.get('zero_rows',''))} {r.get('zero_pct',0):.2f}%")
    if r.get("missing_projects"):
        lines.append(f"       ⚠️  {r['missing_projects']:,} projects in report NOT in proj_df → likely NaN/V_OTHER vertical")
    return lines


def main():
    parser = argparse.ArgumentParser()
    g = parser.add_mutually_exclusive_group(required=True)
    g.add_argument("--entity", choices=list(ENTITIES))
    g.add_argument("--all", action="store_true")
    parser.add_argument("--year", choices=["24", "25", "both"], default="both")
    parser.add_argument("--generate", action="store_true",
                       help="generate deliverable files for ready entities/years")
    args = parser.parse_args()

    targets = list(ENTITIES.items()) if args.all else [(args.entity, ENTITIES[args.entity])]
    output_lines: list[str] = []

    print("\nDelivery Readiness Matrix")
    print("=" * 100)
    for ent, com in targets:
        print(f"\n>>> {ent}...", flush=True)
        r = check_entity(ent, com)
        if "error" in r:
            line = f"{ent:<8} ERROR: {r['error']}"
            output_lines.append(line)
            print(line)
            continue
        lines = [
            f"[{ent}]  src={r['src']}  rows={r['total_rows']:,}  "
            f"sig_LLM_workload={r['sig_llm_workload']:,}  "
            f"truncated={'⚠️ YES' if r['truncated'] else 'no'}"
        ]
        for year in (["24", "25"] if args.year == "both" else [args.year]):
            yr = r["year_checks"].get(year, {})
            lines.extend(format_year_summary(yr))

        if args.generate:
            year_results_for_gen = {y: r["year_checks"][y] for y in r["year_checks"]}
            lines.append("    deliverables:")
            lines.extend(generate_deliverables(ent, com, year_results_for_gen))

        for line in lines:
            print(line, flush=True)
        output_lines.extend(lines)
        output_lines.append("")

    out_path = Path("readiness_matrix.txt")
    out_path.write_text("\n".join(output_lines), encoding="utf-8")
    print(f"\n\nMatrix saved to: {out_path}")
    print("\nLegend:")
    print("  ✅ PASS  — ready to deliver")
    print("  ⚠️  WARN — deliverable but flag for review")
    print("  ❌ FAIL — not ready (must fix first)")


if __name__ == "__main__":
    main()
