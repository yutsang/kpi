"""MGM 25 — aggregate sum BY PROJECT × {Payroll, Capex, Opex} with subtotals.

MGM's raw `wd` column encodes the spend type:
  CAPEX        → Capex
  WD3          → Payroll   (WD3 is the Payroll pivot per conf comment)
  Gaming_OPEX  → Opex
  WD1/WD2/WD4/WD5_Patron → Opex   (operating / COGS / patron comp)

This script reads the post-kedro kpi_report, filters to year 25, maps `wd`
to a 3-bucket spend type, and produces an Excel for manual review:

  data/review/mgm_25_project_payroll_capex_opex.xlsx
    1_投資方向_project   — (投資方向 vertical → project) × {Payroll, Capex, Opex, 合計}
                           with 小計 per 投資方向 + 總計 grand total
    2_project_x_wd_raw   — project × raw wd values (to VERIFY the bucket mapping)
    3_wd_legend          — wd value → bucket + row count + Σ amount (sanity check)

Run:
  python scripts/mgm_project_payroll_capex_opex.py
  python scripts/mgm_project_payroll_capex_opex.py --year 25
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
import pandas as pd
import yaml

COM = "company_6"
ENT = "mgm"
YEAR_CANDIDATES = ("report_period", "report_year", "Yr related", "years")

# wd value → spend bucket. Edit here if the mapping is wrong (user will verify
# via sheet 3_wd_legend). Anything not listed falls to Opex with a warning.
WD_BUCKET = {
    "CAPEX": "Capex",
    "WD3": "Payroll",
    "Gaming_OPEX": "Opex",
    "WD1": "Opex",
    "WD2": "Opex",
    "WD4": "Opex",       # 餐飲收支 COGS
    "WD5_Patron": "Opex",  # patron comp
}
BUCKETS = ["Payroll", "Capex", "Opex"]


def _fuzzy_col(df, name):
    if not name: return None
    if name in df.columns: return name
    for c in df.columns:
        if str(c).strip() == str(name).strip(): return c
    return None


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--year", default="25")
    args = p.parse_args()

    parquet = Path(f"data/{ENT}/output/{COM}_kpi_report.parquet")
    if not parquet.exists():
        print(f"❌ {parquet} missing — run kedro first"); sys.exit(1)

    cfg = yaml.safe_load(Path(f"conf/{COM}/parameters.yml").read_text(encoding="utf-8"))
    cols = cfg.get("columns", {}) or {}

    df = pd.read_parquet(parquet)
    print(f"loaded {len(df):,} rows", flush=True)

    amt_col = _fuzzy_col(df, cols.get("amount", "")) or "amount_mop"
    wd_col = _fuzzy_col(df, cols.get("capex_opex", "")) or "wd"
    proj_col = _fuzzy_col(df, cols.get("project", "")) or "project_name"

    df["amount_mop"] = pd.to_numeric(df[amt_col], errors="coerce").fillna(0)
    df["_wd"] = df[wd_col].astype(str).str.strip() if wd_col in df.columns else ""
    df["_project"] = df[proj_col].astype(str).str.strip() if proj_col in df.columns else ""
    df["_vlabel"] = df["vertical_label"].astype(str) if "vertical_label" in df.columns else ""

    # Filter to year 25
    ycol = next((c for c in YEAR_CANDIDATES if c in df.columns), None)
    if ycol:
        s = df[ycol].astype(str)
        df = df[s.str.startswith(args.year) | (s == f"Yr 20{args.year}") | (s == f"20{args.year}")].copy()
    print(f"year-{args.year} rows: {len(df):,}", flush=True)
    if len(df) == 0:
        print("no rows — abort"); sys.exit(1)

    # Map wd → bucket; warn on unmapped
    seen_wd = sorted(df["_wd"].unique())
    unmapped = [w for w in seen_wd if w not in WD_BUCKET]
    if unmapped:
        print(f"⚠ unmapped wd values (→ Opex by default): {unmapped}", flush=True)
    df["_bucket"] = df["_wd"].map(lambda w: WD_BUCKET.get(w, "Opex"))

    out_dir = Path("data/review")
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"mgm_{args.year}_project_payroll_capex_opex.xlsx"

    with pd.ExcelWriter(out_path, engine="xlsxwriter") as w:
        # ── Sheet 1: 投資方向 → project × {Payroll, Capex, Opex, 合計} with subtotals ──
        piv = df.pivot_table(index=["_vlabel", "_project"], columns="_bucket",
                             values="amount_mop", aggfunc="sum", fill_value=0,
                             observed=True)
        for b in BUCKETS:
            if b not in piv.columns:
                piv[b] = 0.0
        piv = piv[BUCKETS]

        rows = []
        grand = {b: 0.0 for b in BUCKETS}
        for vlabel, sub in piv.groupby(level=0):
            sub_v = sub.droplevel(0)
            for proj, r in sub_v.iterrows():
                rows.append({"投資方向": vlabel, "project": proj,
                             **{b: r[b] for b in BUCKETS},
                             "合計": sum(r[b] for b in BUCKETS)})
            sv = {b: sub_v[b].sum() for b in BUCKETS}
            rows.append({"投資方向": vlabel, "project": f"小計 — {vlabel}",
                         **sv, "合計": sum(sv.values())})
            for b in BUCKETS:
                grand[b] += sv[b]
        rows.append({"投資方向": "總計", "project": "總計",
                     **grand, "合計": sum(grand.values())})
        out1 = pd.DataFrame(rows, columns=["投資方向", "project", *BUCKETS, "合計"])
        out1.to_excel(w, sheet_name="1_投資方向_project", index=False)
        print(f"sheet 1: {len(out1):,} rows ({piv.shape[0]} projects + subtotals)", flush=True)

        # ── Sheet 2: project × raw wd (verification) ──
        piv2 = df.pivot_table(index="_project", columns="_wd",
                              values="amount_mop", aggfunc="sum", fill_value=0,
                              observed=True, margins=True, margins_name="總計")
        piv2.to_excel(w, sheet_name="2_project_x_wd_raw")
        print(f"sheet 2: {piv2.shape[0]} project rows × {piv2.shape[1]} wd cols", flush=True)

        # ── Sheet 3: wd legend (mapping + counts + Σ amount) ──
        leg = df.groupby("_wd", observed=True).agg(
            bucket=("_bucket", "first"),
            rows=("amount_mop", "size"),
            amount_mop=("amount_mop", "sum"),
        ).reset_index().rename(columns={"_wd": "wd_value"}).sort_values("amount_mop", ascending=False)
        leg.to_excel(w, sheet_name="3_wd_legend", index=False)
        print(f"sheet 3: {len(leg)} distinct wd values", flush=True)

    size_mb = out_path.stat().st_size / 1e6
    print(f"\n✓ {out_path}  ({size_mb:.1f} MB)", flush=True)
    print(f"\nGrand total: Payroll={grand['Payroll']/1e6:,.1f}M  "
          f"Capex={grand['Capex']/1e6:,.1f}M  Opex={grand['Opex']/1e6:,.1f}M  "
          f"合計={sum(grand.values())/1e6:,.1f}M", flush=True)
    print("\n下一步: 開 sheet 3_wd_legend 確認 wd→bucket 對唔對 (尤其 WD1/WD2/WD3/WD4)，"
          "如果 mapping 錯 → 改 WD_BUCKET dict 再跑。", flush=True)


if __name__ == "__main__":
    main()
