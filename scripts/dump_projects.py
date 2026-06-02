"""Dump distinct projects per entity (1 row each) with current V + the project-team's own
category column(s) + Σamount + row count — to RE-CHECK V labels after a taxonomy change,
without re-reading every JE line. Sorted by |Σamount| so the money-relevant ones are on top.

Use after kedro: spot projects whose current V looks wrong under the new taxonomy (e.g. should
be 節日慶典 / 內部設施升級 / 公共設施設備提升), then I add targeted row_vertical_overrides.

Run:
  python scripts/dump_projects.py --entity galaxy --year 25 --n 80
  python scripts/dump_projects.py --entity vml --year 25 --v 其他      # only V=其他 rows (find mis-tags)
  python scripts/dump_projects.py            # all 6 entities, top 60 each, full → results/
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
# project-team category columns that drive / hint V (whichever exist get shown)
CAT_COLS = ["分類1", "項目性質", "項目分類2", "範疇", "範疇第二層標籤", "NG11 Category",
            "Project Session", "一級標籤", "二級標籤", "一級標簽", "二級標簽", "投資方向",
            "類別", "comp费用大类", "Nature of Expenses"]


def _col(df, *cands):
    for c in cands:
        if c and c in df.columns:
            return c
    return None


def dump(ent, com, args):
    pq = ROOT / f"data/{ent}/output/{com}_kpi_report.parquet"
    if not pq.exists():
        print(f"[{ent}] {pq.name} missing — skip"); return
    cfg = yaml.safe_load((ROOT / f"conf/{com}/parameters.yml").read_text(encoding="utf-8"))
    cols = cfg.get("columns", {})
    # Read ONLY needed columns + strip pandas StringDtype metadata: avoids the
    # ArrowStringArray.__from_arrow__ crash on some unicode AND lowers memory (tight box).
    import pyarrow.parquet as _pq
    _names = _pq.read_schema(pq).names
    _want = [c for c in ("report_period", "report_year", "years",
                         "project", cols.get("project"), "SubProject_Name", "Name of Investment Project",
                         cols.get("amount"), "amount_mop", "amount",
                         "vertical_id", "vertical_label", *CAT_COLS) if c and c in _names]
    _seen = set(); _want = [c for c in _want if not (c in _seen or _seen.add(c))]
    df = _pq.read_table(pq, columns=_want).replace_schema_metadata(None).to_pandas()
    yc = _col(df, "report_period", "report_year", "years")
    if yc:
        df = df[df[yc].astype(str).str.startswith(args.year)]
    proj = _col(df, "project", cols.get("project"), "SubProject_Name", "Name of Investment Project")
    amtc = _col(df, cols.get("amount"), "amount_mop", "amount")
    vi = _col(df, "vertical_id"); vl = _col(df, "vertical_label")
    if not proj or not vi:
        print(f"[{ent}] no project/vertical cols — skip"); return
    df["_amt"] = pd.to_numeric(df[amtc], errors="coerce").fillna(0) if amtc else 0.0
    if args.v and vl:
        df = df[df[vl].astype(str).str.contains(args.v, na=False) | df[vi].astype(str).str.contains(args.v, na=False)]
    cats = [c for c in CAT_COLS if c in df.columns]

    def _mode(s):
        s = s[s.astype(str).str.strip().ne("")]
        return s.value_counts().index[0] if len(s) else ""
    agg = {"_amt": ("_amt", "sum"), "rows": ("_amt", "size"), "cur_V": (vi, _mode)}
    if vl:
        agg["cur_Vlabel"] = (vl, _mode)
    for c in cats:
        agg[c] = (c, _mode)
    g = df.groupby(df[proj].astype(str)).agg(**agg).reset_index().rename(columns={proj: "project"})
    g = g.reindex(g["_amt"].abs().sort_values(ascending=False).index).reset_index(drop=True)
    g = g.rename(columns={"_amt": "Σamt"})

    show = ["project", "rows", "Σamt", "cur_Vlabel" if vl else "cur_V"] + cats
    show = [c for c in show if c in g.columns]
    print(f"\n=== {ent} {args.year} — {len(g)} distinct projects (top {args.n} by |Σamt|; cats: {cats or 'none'}) ===")
    with pd.option_context("display.max_colwidth", 42, "display.width", 200):
        print(g.head(args.n)[show].to_string(index=False))
    rep = ROOT / "results" / f"projects_{ent}_{args.year}.tsv"
    rep.parent.mkdir(parents=True, exist_ok=True)
    g.to_csv(rep, sep="\t", index=False, encoding="utf-8-sig")
    print(f"→ full {len(g)} projects: {rep}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", choices=list(ENTITIES), default=None)
    ap.add_argument("--year", default="25")
    ap.add_argument("--v", default=None, help="filter to current V (label/id substring)")
    ap.add_argument("--n", type=int, default=60)
    args = ap.parse_args()
    targets = [(args.entity, ENTITIES[args.entity])] if args.entity else list(ENTITIES.items())
    for ent, com in targets:
        dump(ent, com, args)


if __name__ == "__main__":
    main()
