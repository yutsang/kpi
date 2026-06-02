"""Find the project-team MANUAL classification columns in an entity's tagged_rows.

Every entity's raw data carries the project team's own labels somewhere (a 分類/
類別/性質/標籤/領域 column, or comp-nature, etc.). Those are the ground truth we
map to OUR taxonomy. This scans tagged_rows, flags every column that LOOKS like a
manual category (low-cardinality + CJK / name match), and dumps each candidate's
value → row-count → Σ|amount| so we can build the their-category → our V/H map.

Outputs (results/):
  {ent}_manual_cols.tsv    — every flagged candidate column × value × n × amount
  console: candidate column list + each candidate's value distribution

Run:
  python scripts/inspect_manual_cols.py --entity wynn
  python scripts/inspect_manual_cols.py --entity galaxy --year 25
  python scripts/inspect_manual_cols.py --entity melco
"""
from __future__ import annotations
import argparse, sys, csv, re
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
import pandas as pd
import yaml

ENT = {"galaxy":"company_1","sjm":"company_2","wynn":"company_3",
       "vml":"company_4","melco":"company_5","mgm":"company_6"}
NAME_HINT = re.compile(r"分類|分类|類別|类别|性質|性质|標籤|标签|領域|领域|level|category|"
                       r"nature|一級|二級|一级|二级|comp|投資方向|投资方向|類型|类型")
CJK = re.compile(r"[一-鿿]")


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--entity", required=True, choices=list(ENT))
    p.add_argument("--year", default=None)
    p.add_argument("--min_uniq", type=int, default=2)
    p.add_argument("--max_uniq", type=int, default=80)
    p.add_argument("--raw", action="store_true",
                   help="read the raw Excel sheet (from conf yearly_sources) instead of tagged_rows")
    p.add_argument("--all", action="store_true",
                   help="dump EVERY column (name + cardinality + samples), not just flagged candidates")
    args = p.parse_args()
    ent = args.entity; com = ENT[ent]

    cfg = yaml.safe_load(Path(f"conf/{com}/parameters.yml").read_text(encoding="utf-8"))
    cols = (cfg.get("columns") or {})

    if args.raw:
        # pick the raw file+sheet from yearly_sources (match --year, else first / legacy)
        ys = cfg.get("yearly_sources") or []
        src = None
        if args.year:
            src = next((s for s in ys if str(s.get("year", "")).startswith(args.year.rstrip("年度"))
                        or str(s.get("year", ""))[-2:] == args.year[-2:]), None)
        src = src or (ys[0] if ys else {})
        rawfn = src.get("raw_filename") or cfg.get("raw_filename")
        sheet = src.get("raw_sheet", cfg.get("raw_sheet", 0))
        hdr = cfg.get("header_row", 0)
        rawpath = Path(f"data/{ent}/raw/{rawfn}")
        if not rawpath.exists():
            print(f"X {rawpath} missing"); sys.exit(1)
        print(f"[{ent}] reading RAW {rawpath.name}::{sheet} (header_row={hdr})")
        df = pd.read_excel(rawpath, sheet_name=sheet, header=hdr, dtype=str)
    else:
        pq = Path(f"data/{ent}/interim/{com}_tagged_rows.parquet")
        if not pq.exists():
            print(f"X {pq} missing — run kedro first (or use --raw to read the source Excel)"); sys.exit(1)
        df = pd.read_parquet(pq)
    amt = next((c for c in (cols.get("amount"), "amount_mop", "MOP Amt", "original_amount")
                if c and c in df.columns), None)
    if args.year:
        rp = next((c for c in ("report_period", "report_year") if c in df.columns), None)
        if rp:
            df = df[df[rp].astype(str).str.strip().str.startswith(args.year)].copy()
    amt_s = pd.to_numeric(df[amt], errors="coerce").fillna(0) if amt else pd.Series(0, index=df.index)
    print(f"[{ent}] {len(df):,} rows | amt={amt} | year={args.year or 'all'}")

    # ---- --all: dump EVERY column (so we never miss a label column) ----
    if args.all:
        out = Path("results"); out.mkdir(exist_ok=True)
        arows = []
        print(f"\n=== ALL {len(df.columns)} columns ===")
        for c in df.columns:
            try:
                s = df[c].astype("string").fillna("").str.strip()
                nb = s[s.ne("") & s.ne("nan") & s.ne("<NA>")]
                n = nb.nunique()
                samp = " | ".join(str(x)[:30] for x in nb.unique()[:6])
            except Exception:
                n, samp = -1, ""
            cjk = "CJK" if CJK.search(samp or "") else "   "
            print(f"  {cjk} {str(c)[:40]:<40} n_distinct={n:<6} {samp[:70]}")
            arows.append([str(c), n, samp])
        with (out/f"{ent}_all_cols.tsv").open("w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter="\t")
            w.writerow(["column", "n_distinct", "sample_values"])
            w.writerows(arows)
        print(f"\n→ results/{ent}_all_cols.tsv")

    # ---- flag candidate manual-category columns ----
    cand = []
    for c in df.columns:
        try:
            s = df[c].astype("string").fillna("").str.strip()
            nonblank = s[s.ne("") & s.ne("nan") & s.ne("<NA>")]
            n = nonblank.nunique()
        except Exception:
            continue
        if n < args.min_uniq or n > args.max_uniq:
            continue
        sample = nonblank.unique()[:5]
        has_cjk = any(CJK.search(str(x)) for x in sample)
        name_hit = bool(NAME_HINT.search(str(c)))
        # candidate if name hints OR (low-card + mostly CJK values + not a tag col we made)
        if str(c) in ("vertical_id","vertical_label","horizontal_id","horizontal_label",
                      "vertical_source","horizontal_source","ng_scope","row_type",
                      "final_capex_opex"):
            continue
        if name_hit or (has_cjk and n <= 40):
            cand.append((c, n, name_hit, " | ".join(str(x)[:24] for x in sample)))
    cand.sort(key=lambda r: (not r[2], r[1]))

    print(f"\n=== {len(cand)} candidate manual-category columns ===")
    for c, n, hit, samp in cand:
        print(f"  {'★' if hit else ' '} {str(c)[:34]:<34} n={n:<4} {samp[:60]}")

    # ---- per-candidate value distribution ----
    out = Path("results"); out.mkdir(exist_ok=True)
    rows = []
    for c, n, hit, _ in cand:
        s = df[c].astype("string").fillna("").str.strip()
        g = pd.DataFrame({"v": s, "a": amt_s}).groupby("v", dropna=False)["a"].agg(
            n="size", amount=lambda x: x.abs().sum()).reset_index()
        g = g[g["v"].ne("") & g["v"].ne("nan") & g["v"].ne("<NA>")]
        g = g.sort_values("amount", ascending=False)
        for _, r in g.iterrows():
            rows.append([c, r["v"], int(r["n"]), round(r["amount"], 0)])
    with (out/f"{ent}_manual_cols.tsv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter="\t")
        w.writerow(["column", "value", "n_rows", "amount"])
        w.writerows(rows)

    # console: top candidates' value dist (the ★ name-hit ones)
    for c, n, hit, _ in [x for x in cand if x[2]][:6]:
        s = df[c].astype("string").fillna("").str.strip()
        g = pd.DataFrame({"v": s, "a": amt_s}).groupby("v")["a"].agg(
            n="size", amount=lambda x: x.abs().sum()).reset_index()
        g = g[g["v"].ne("") & g["v"].ne("nan") & g["v"].ne("<NA>")].sort_values("amount", ascending=False)
        print(f"\n--- {c} ({len(g)} values) ---")
        for _, r in g.head(30).iterrows():
            print(f"    {str(r['v'])[:40]:<40} n={int(r['n']):>6} amt={r['amount']/1e6:>10,.2f}M")
    print(f"\n→ results/{ent}_manual_cols.tsv")


if __name__ == "__main__":
    main()
