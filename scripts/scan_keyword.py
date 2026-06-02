"""Scan all 6 entities' kpi_report for keyword hits — enumerate candidate projects/rows for a
new vertical (e.g. wifi/系統 → 內部設施升級) so we route the RIGHT ones instead of guessing.

For each keyword, per entity: #rows, Σamount, and a few distinct project/description samples.
Run AFTER kedro (needs data/{ent}/output/{com}_kpi_report.parquet).

Run:
  python scripts/scan_keyword.py
  python scripts/scan_keyword.py --keywords "wifi,系統,網絡,IT,資訊,server,智慧,digital,基礎設施"
  python scripts/scan_keyword.py --year 25
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parent.parent
ENTITIES = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
            "vml": "company_4", "melco": "company_5", "mgm": "company_6"}

DEFAULT_KW = ["wifi", "wi-fi", "系統", "網絡", "網路", "IT", "資訊科技", "資訊系統",
              "network", "system", "server", "伺服器", "軟件", "software", "智慧",
              "smart", "數字化", "digital", "基礎設施", "infrastructure", "電腦"]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--keywords", default=",".join(DEFAULT_KW))
    ap.add_argument("--year", default="25")
    ap.add_argument("--samples", type=int, default=4)
    args = ap.parse_args()
    kws = [k.strip() for k in args.keywords.split(",") if k.strip()]

    out_rows = []
    for ent, com in ENTITIES.items():
        pq = ROOT / f"data/{ent}/output/{com}_kpi_report.parquet"
        if not pq.exists():
            print(f"[{ent}] {pq.name} missing — skip"); continue
        cfg = yaml.safe_load((ROOT / f"conf/{com}/parameters.yml").read_text(encoding="utf-8"))
        cols = cfg.get("columns", {})
        import pyarrow.parquet as _pq  # strip pandas StringDtype metadata → object (avoids __from_arrow__ crash)
        df = _pq.read_table(pq).replace_schema_metadata(None).to_pandas()
        yc = next((c for c in ("report_period", "report_year", "years") if c in df.columns), None)
        if yc:
            df = df[df[yc].astype(str).str.startswith(args.year)]
        # search across project / subproject / description / vendor + their conf names
        scan_cols = []
        for key in ("project", "description", "vendor"):
            c = cols.get(key)
            if c and c in df.columns:
                scan_cols.append(c)
        for c in ("project", "subproject", "description", "Name of Investment Project"):
            if c in df.columns and c not in scan_cols:
                scan_cols.append(c)
        if not scan_cols:
            print(f"[{ent}] no text cols"); continue
        amt = pd.to_numeric(df[cols.get("amount", "amount_mop")], errors="coerce").fillna(0) \
            if cols.get("amount", "amount_mop") in df.columns else pd.Series(0.0, index=df.index)
        hay = df[scan_cols[0]].astype(str)
        for c in scan_cols[1:]:
            hay = hay + " ¦ " + df[c].astype(str)
        hay_l = hay.str.lower()
        print(f"\n=== {ent} ({len(df):,} rows, year {args.year}; cols={scan_cols}) ===")
        for kw in kws:
            m = hay_l.str.contains(kw.lower(), na=False, regex=False)
            n = int(m.sum())
            if n == 0:
                continue
            s = float(amt[m].sum())
            samples = hay[m].drop_duplicates().head(args.samples).tolist()
            print(f"  {kw:14} rows={n:>6}  Σ={s:>16,.0f}")
            for sp in samples:
                print(f"       · {str(sp)[:90]}")
            out_rows.append({"entity": ent, "keyword": kw, "rows": n, "amount": round(s)})

    if out_rows:
        rep = ROOT / "results" / "keyword_scan.tsv"
        rep.parent.mkdir(parents=True, exist_ok=True)
        pd.DataFrame(out_rows).to_csv(rep, sep="\t", index=False, encoding="utf-8-sig")
        print(f"\n→ {rep}")


if __name__ == "__main__":
    main()
