"""inspect_vproj.py — per-PROJECT vertical (V) comparison: OUR vertical_label vs 項目組's own
project label (類別2 / 類別1) for a year. Taxonomy differs, but obvious mis-V'd projects pop out.
Output → paste → Claude flags the obvious ones → row_vertical_overrides / manual_vertical fix.

  python scripts/inspect_vproj.py --entity vml --year 23
  python scripts/inspect_vproj.py --entity vml --year 23 --top all
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
ENT = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
       "vml": "company_4", "melco": "company_5", "mgm": "company_6"}


def numify(s):
    return pd.to_numeric(s.astype(str).str.replace(",", "", regex=False).str.replace(r"^\s*-\s*$", "0", regex=True),
                         errors="coerce").fillna(0)


def find(df, *subs, exact=None):
    if exact and exact in df.columns:
        return exact
    for c in df.columns:
        if any(s in str(c) for s in subs):
            return c
    return None


def mode(s):
    s = s.astype(str).str.strip()
    s = s[s.ne("") & s.ne("nan")]
    return s.mode().iloc[0] if len(s.mode()) else ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", required=True, choices=sorted(ENT))
    ap.add_argument("--year", required=True)
    ap.add_argument("--theircol", default=None, help="項目組 V column (default 類別2)")
    ap.add_argument("--top", default="all")
    a = ap.parse_args()
    com = ENT[a.entity]
    src = ROOT / "data" / a.entity / "interim" / f"{com}_tagged_rows.parquet"
    if not src.exists():
        print(f"X missing {src.relative_to(ROOT)} — run kedro."); return
    cf = yaml.safe_load((ROOT / "conf" / com / "parameters.yml").open(encoding="utf-8"))
    cols = cf.get("columns", {}) or {}
    df = pq.read_table(src).replace_schema_metadata(None).to_pandas()
    if "report_period" in df.columns:
        df = df[df["report_period"].astype(str).str.startswith(a.year)].copy()

    amt = next((c for c in [cols.get("amount"), "MOP Amt", "調整後金額", "Reported Amount(MOP)"]
                if c and c in df.columns and numify(df[c]).abs().sum() > 0), None)
    df["_amt"] = numify(df[amt]) if amt else 0.0
    proj = find(df, exact=cols.get("project")) or find(df, "SubProject_Name", "Project")
    ourv = "vertical_label" if "vertical_label" in df.columns else "vertical_id"
    their = a.theircol or find(df, exact="類別2") or find(df, "類別2", "类别2", "類別1", "类别1")
    print(f"[{a.entity} {a.year}] project={proj!r}  our_V={ourv!r}  項目組_V={their!r}  amount={amt!r}")
    if not (proj and their):
        print("  ⚠ need project + 項目組 V col — check column names"); return

    df["_p"] = df[proj].astype(str).str.strip()
    rows = []
    for p, g in df.groupby("_p"):
        rows.append({"project": p, "our_V": mode(g[ourv]),
                     "their_V": mode(g[their]), "n_their": g[their].astype(str).str.strip().replace("nan", "").ne("").sum() and g[their].nunique(),
                     "amt": g["_amt"].sum()})
    out = pd.DataFrame(rows)
    out = out.reindex(out["amt"].abs().sort_values(ascending=False).index)
    n = len(out) if a.top == "all" else int(a.top)
    print(f"\n=== per-project: our_V vs 項目組 類別2  ({len(out)} projects, showing {n}) ===")
    print("  project | our_V | 項目組_類別2 | Σamt")
    for _, r in out.head(n).iterrows():
        flag = ""  # raw — Claude eyeballs the mismatches
        print(f"  {str(r['project'])[:34]:36s} {str(r['our_V'])[:14]:16s} {str(r['their_V'])[:18]:20s} {r['amt']:>15,.0f}{flag}")


if __name__ == "__main__":
    main()
