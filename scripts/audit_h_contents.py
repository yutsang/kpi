"""Per-bucket CONTENT audit — "睇每個 H / V / NG label 入面實際裝住啲咩,有冇唔應該出現嘅".

--by H  : per horizontal_label → top account_code|account_desc inside (is 餐飲 all F&B?)
--by V  : per vertical_label   → top account|desc inside (is V_CONCERT spend sensible?)
--by NG : per NG bucket        → top account|desc inside (is NG3 all entertainment-nature?)

Also flags the real tell of a mis-assignment (only for --by H): an account_desc that lands
in 2+ different H labels with substantial $ in each (same account split = inconsistent rule).

  python scripts/audit_h_contents.py --entity galaxy --year 25            # default --by H
  python scripts/audit_h_contents.py --entity galaxy --year 25 --by V
  python scripts/audit_h_contents.py --entity mgm    --year 25 --by NG --topn 12

Output: console table + results/{ent}_contents_{by}_{year}.tsv  (paste the .tsv).
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import pyarrow.parquet as pq
import yaml

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))
ENTITIES = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
            "vml": "company_4", "melco": "company_5", "mgm": "company_6"}


def fuzzy(df, name):
    if not name:
        return None
    if name in df.columns:
        return name
    for c in df.columns:
        if str(c).strip() == str(name).strip():
            return c
    return None


def audit(ent, year, topn, by):
    com = ENTITIES[ent]
    p = ROOT / "data" / ent / "interim" / f"{com}_tagged_rows.parquet"
    if not p.exists():
        print(f"[{ent}] X missing {p.relative_to(ROOT)}"); return
    cf = yaml.safe_load((ROOT / "conf" / com / "parameters.yml").open(encoding="utf-8"))
    cols = cf.get("columns", {}) or {}
    df = pq.read_table(p).replace_schema_metadata(None).to_pandas()
    ac = fuzzy(df, cols.get("account_code"))
    ad = fuzzy(df, cols.get("account_desc"))
    am = fuzzy(df, cols.get("amount"))
    if not am:
        print(f"[{ent}] X need amount"); return
    df["_amt"] = pd.to_numeric(df[am], errors="coerce").fillna(0.0)
    df["_acc"] = ((df[ac].astype(str) + "|") if ac else "") + (df[ad].astype(str) if ad else "")

    if by == "V":
        if "vertical_label" not in df.columns:
            print(f"[{ent}] X no vertical_label"); return
        df["_G"] = df["vertical_label"].astype(str)
    elif by == "NG":
        from audit_health import derive_ng
        df["_G"] = derive_ng(df, cf, cols).astype(str)
    else:
        if "horizontal_label" not in df.columns:
            print(f"[{ent}] X no horizontal_label"); return
        df["_G"] = df["horizontal_label"].astype(str)

    if year and "report_period" in df.columns:
        df = df[df["report_period"].astype(str).str.startswith(str(year))].copy()
    if df.empty:
        print(f"[{ent} {year}] no rows"); return

    grand = df["_amt"].sum()
    print(f"\n##########  {ent.upper()} {year or 'ALL'}  by {by}  Σ={grand:,.0f}  rows={len(df):,}  ##########")

    # ── cross-bucket split (most meaningful for H) ──
    acc_h = df.groupby(["_acc", "_G"])["_amt"].agg(lambda s: s.abs().sum()).reset_index()
    split_rows = []
    for acc, g in acc_h.groupby("_acc"):
        tot = g["_amt"].sum()
        if tot <= 0:
            continue
        big = g[g["_amt"] / tot >= 0.05].sort_values("_amt", ascending=False)
        if len(big) >= 2:
            split_rows.append((acc, tot, "  ".join(f"{r['_G']}:{r['_amt']/tot*100:.0f}%" for _, r in big.iterrows())))
    split_rows.sort(key=lambda x: -x[1])

    # ── per-bucket contents ──
    hsum = df.groupby("_G")["_amt"].agg(total="sum", absamt=lambda s: s.abs().sum(), n="size").reset_index()
    hsum = hsum.sort_values("absamt", ascending=False)
    res = ROOT / "results"; res.mkdir(exist_ok=True)
    lines = []
    for _, hr in hsum.iterrows():
        G = hr["_G"]
        g = df[df["_G"] == G]
        sub = (g.groupby("_acc")["_amt"]
                 .agg(amt="sum", absamt=lambda s: s.abs().sum(), n="size")
                 .reset_index().sort_values("absamt", ascending=False).head(topn))
        pct = hr["total"] / grand * 100 if grand else 0
        print(f"\n=== {G}   Σ={hr['total']:,.0f}  ({pct:.1f}% of year)  rows={int(hr['n']):,} ===")
        for _, r in sub.iterrows():
            sharepct = r["absamt"] / hr["absamt"] * 100 if hr["absamt"] else 0
            mark = "  ⚠SPLIT" if any(r["_acc"] == s[0] for s in split_rows) else ""
            print(f"   {r['amt']:>16,.0f}  {sharepct:4.0f}%  n={int(r['n']):>6}  {str(r['_acc'])[:60]}{mark}")
            lines.append({"bucket": G, "account": r["_acc"], "amount": r["amt"], "share_pct": round(sharepct, 1), "n": int(r["n"])})

    if by == "H":
        print(f"\n=== ⚠ account_desc SPLIT across 2+ H buckets (>=5% each) — likely mis-assignments ===")
        for acc, tot, brk in split_rows[:40]:
            print(f"   |Σ|={tot:>15,.0f}  {str(acc)[:46]:48s} {brk}")
        if not split_rows:
            print("   (none — every account sits in a single H bucket)")

    pd.DataFrame(lines).to_csv(res / f"{ent}_contents_{by}_{year or 'all'}.tsv", sep="\t", index=False, encoding="utf-8-sig")
    print(f"\n  ► paste: results/{ent}_contents_{by}_{year or 'all'}.tsv")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", required=True, choices=list(ENTITIES))
    ap.add_argument("--year", default="")
    ap.add_argument("--by", default="H", choices=["H", "V", "NG"])
    ap.add_argument("--topn", type=int, default=10)
    a = ap.parse_args()
    audit(a.entity, a.year, a.topn, a.by)
