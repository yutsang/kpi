"""For an entity+year, compare EACH project-team raw category column (分類1 / 分類2 /
會計科目分類 …) against our pipeline V (vertical_label) and H (horizontal_label) — to see
which raw column is the V source, which is the H source, and whether our result lines up.

Reads {ent} tagged_rows.parquet (keeps ALL raw columns + our V/H, already row-matched — no join
needed). Concentration = for each raw-column value, what % of its amount lands in ONE of our
labels. High top1% = clean 1:1 (just a different name); low = scattered (the problem area).

Run on Windows:
  python scripts/compare_raw_cols.py --entity vml --year 24 --amount 調整後金額
  python scripts/compare_raw_cols.py --entity vml --year 24 --cols "分類1,分類2,會計科目分類"
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
CATPAT = re.compile(r"分類|科目|類別|標簽|標籤|範疇|性質")


def _conc(df, src, dst, amt):
    g = df.groupby([df[src].astype(str), df[dst].astype(str)])[amt].sum().abs().reset_index()
    g.columns = [src, dst, "amt"]
    rows, clean, tot_all = [], 0, 0.0
    for sv, sub in g.groupby(src):
        t = sub["amt"].sum()
        if t <= 0 or str(sv).strip() in ("", "nan", "None"):
            continue
        sub = sub.sort_values("amt", ascending=False)
        top1 = sub.iloc[0]["amt"] / t * 100
        clean += int(top1 >= 80)
        tot_all += t
        rows.append((sv, t, top1, sub[dst].nunique(),
                     " ; ".join(f"{r[dst]}={r['amt']/t*100:.0f}%" for _, r in sub.head(3).iterrows())))
    rows.sort(key=lambda r: -abs(r[1]))
    return rows, clean, tot_all


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", required=True, choices=list(ENTITIES))
    ap.add_argument("--year", default="24")
    ap.add_argument("--amount", default=None, help="amount col (default: 調整後金額 → conf amount → fallback)")
    ap.add_argument("--cols", default=None, help="comma list of raw cols to test (default: auto-detect 分類/科目)")
    ap.add_argument("--n", type=int, default=20)
    args = ap.parse_args()
    com = ENTITIES[args.entity]
    src = ROOT / f"data/{args.entity}/interim/{com}_tagged_rows.parquet"
    if not src.exists():
        print(f"X {src} missing — run kedro {args.entity} first"); return
    cfg = yaml.safe_load((ROOT / f"conf/{com}/parameters.yml").read_text(encoding="utf-8"))
    conf_amt = (cfg.get("columns", {}) or {}).get("amount")

    names = pq.read_schema(src).names
    catcols = ([c.strip() for c in args.cols.split(",")] if args.cols
               else [c for c in names if CATPAT.search(str(c))])
    catcols = [c for c in catcols if c in names]
    amt = next((c for c in (args.amount, "調整後金額", conf_amt, "MOP Amt", "amount_mop", "amount")
                if c and c in names), None)
    keep = list(dict.fromkeys(catcols + [c for c in ("vertical_label", "horizontal_label",
                                                     "report_period", amt) if c and c in names]))
    df = pq.read_table(src, columns=keep).replace_schema_metadata(None).to_pandas()
    if "report_period" in df.columns:
        df = df[df["report_period"].astype(str).str.startswith(args.year)]
    df[amt] = pd.to_numeric(df[amt], errors="coerce").fillna(0) if amt else 1.0

    print(f"=== {args.entity} {args.year} — 原表 columns vs 我們 V/H ===")
    print(f"amount = {amt!r}  rows={len(df):,}  Σ={df[amt].sum():,.0f}")
    print(f"raw category columns tested: {catcols}\n")

    for c in catcols:
        line = f"--- '{c}' ---"
        for axis, dst in (("→我們V", "vertical_label"), ("→我們H", "horizontal_label")):
            if dst not in df.columns:
                continue
            rows, clean, tot = _conc(df, c, dst, amt)
            if not rows:
                continue
            avg = sum(r[2] for r in rows) / len(rows)
            print(f"{line}  {axis}: {len(rows)} 值, 乾淨(top1≥80%)={clean}/{len(rows)}, avg top1%={avg:.0f}%, Σ={tot:,.0f}")
            line = " " * len(f"--- '{c}' ---")
        # show the better-aligned axis detail (V vs H by avg top1%)
        best = max((("vertical_label", "→V"), ("horizontal_label", "→H")),
                   key=lambda kv: (sum(r[2] for r in _conc(df, c, kv[0], amt)[0]) /
                                   max(1, len(_conc(df, c, kv[0], amt)[0]))) if kv[0] in df.columns else 0)
        rows, _, _ = _conc(df, c, best[0], amt)
        print(f"   top {args.n} ({c} {best[1]}):")
        for sv, t, top1, nd, tops in rows[:args.n]:
            flag = "  ⚠" if top1 < 80 else ""
            print(f"     {str(sv)[:22]:22} Σ={t:>14,.0f} top1={top1:>5.0f}% nd={nd:>2}  {tops[:62]}{flag}")
        print()


if __name__ == "__main__":
    main()
