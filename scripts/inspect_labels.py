"""Inspect the project-team's own label columns (一級/二級標簽 + 人手/人工 manual) behind a
given V/H — to fix H mis-mappings (e.g. Galaxy 娛樂表演→活動場地 450M: which 二級標簽 / 人手
values feed it, so we re-map the column_map).

Reads {ent} tagged_rows.parquet (it keeps ALL raw columns incl the label ones, unlike
kpi_report). Lean: reads only the detected label cols + V/H + amount (+ strips pandas
StringDtype metadata to dodge the arrow unicode crash + lower RAM).

Run:
  python scripts/inspect_labels.py --entity galaxy --year 25 --v 娛樂表演 --h 活動場地
  python scripts/inspect_labels.py --entity galaxy --year 25 --h 活動場地     # all V
"""
from __future__ import annotations
import argparse, re, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import pyarrow.parquet as pq
import yaml

ROOT = Path(__file__).resolve().parent.parent
ENTITIES = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
            "vml": "company_4", "melco": "company_5", "mgm": "company_6"}
LABEL_PAT = re.compile(r"一級|二級|三級|人手|人工|標簽|標籤|分類|範疇|性質|費用大类|level|categor|nature", re.I)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", required=True, choices=list(ENTITIES))
    ap.add_argument("--year", default="25")
    ap.add_argument("--v", default=None, help="filter current vertical_label (substring)")
    ap.add_argument("--h", default=None, help="filter current horizontal_label (substring)")
    ap.add_argument("--n", type=int, default=40)
    args = ap.parse_args()
    com = ENTITIES[args.entity]
    src = ROOT / f"data/{args.entity}/interim/{com}_tagged_rows.parquet"
    if not src.exists():
        print(f"X {src} missing — run kedro {args.entity} step4 first"); return
    cfg = yaml.safe_load((ROOT / f"conf/{com}/parameters.yml").read_text(encoding="utf-8"))
    cols = cfg.get("columns", {})

    names = pq.read_schema(src).names
    label_cols = [c for c in names if LABEL_PAT.search(str(c))]
    keep = label_cols + [c for c in ("report_period", "vertical_label", "vertical_id",
                                     "horizontal_label", "horizontal_id",
                                     cols.get("account_code"), cols.get("account_desc"),
                                     cols.get("project"), "project",
                                     cols.get("amount"), "amount_mop") if c and c in names]
    seen = set(); keep = [c for c in keep if not (c in seen or seen.add(c))]
    df = pq.read_table(src, columns=keep).replace_schema_metadata(None).to_pandas()

    yc = next((c for c in ("report_period",) if c in df.columns), None)
    if yc:
        df = df[df[yc].astype(str).str.startswith(args.year)]
    if args.v and "vertical_label" in df.columns:
        df = df[df["vertical_label"].astype(str).str.contains(args.v, na=False)]
    if args.h and "horizontal_label" in df.columns:
        df = df[df["horizontal_label"].astype(str).str.contains(args.h, na=False)]
    amtc = next((c for c in (cols.get("amount"), "amount_mop", "amount") if c and c in df.columns), None)
    df["_amt"] = pd.to_numeric(df[amtc], errors="coerce").fillna(0) if amtc else 0.0
    print(f"=== {args.entity} {args.year} | V~{args.v} H~{args.h} | {len(df):,} rows  Σ={df['_amt'].sum():,.0f} ===")
    print(f"label columns detected: {label_cols}")
    if df.empty:
        return

    gcols = [c for c in label_cols if c in df.columns]
    if gcols:
        def _mode(s):
            s = s.astype(str)
            return s.mode().iloc[0] if len(s) else ""
        g = df.groupby([df[c].astype(str) for c in gcols]).agg(
            cur_V=("vertical_label", _mode), cur_H=("horizontal_label", _mode),
            Σamt=("_amt", "sum"), rows=("_amt", "size")).reset_index()
        g = g.reindex(g["Σamt"].abs().sort_values(ascending=False).index)
        print(f"\n— by label-combo ({gcols}) → current H, top {args.n} —")
        with pd.option_context("display.max_colwidth", 30, "display.width", 220):
            print(g.head(args.n).to_string(index=False))
        rep = ROOT / "results" / f"labels_{args.entity}_{args.year}.tsv"
        rep.parent.mkdir(parents=True, exist_ok=True)
        g.to_csv(rep, sep="\t", index=False, encoding="utf-8-sig")
        print(f"→ {rep}")


if __name__ == "__main__":
    main()
