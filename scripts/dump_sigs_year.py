"""Dump ONE report-year's unique signatures (from tagged_rows) for H review / redo.

`unique_signatures.xlsx` is cross-year; this reads the per-row `tagged_rows.parquet`
so we can scope to a single report year (e.g. VML 2023) and see, per signature:
  current_H  — the FINAL H on those rows (horizontal_id, after step4)
  peer_H     — the H that already-classified sigs of the SAME account_desc got across
               ALL years (a consistency hint: usually current should follow peer)
  account_code/desc, desc_sample, n_rows, amount, projects

  python scripts/dump_sigs_year.py --entity vml --year 23                 # all 2023 sigs
  python scripts/dump_sigs_year.py --entity vml --year 23 --unclassified  # only H_OTHER (the gap)

Output: data/review/_dump/{ent}_sigs_{year}[_unclassified].tsv  (+ ≤3000-row batches)
Then put the TSV in results/ — Claude classifies H per signature → {ent}_horizontal.tsv,
inject_horizontal_feedback applies it (signature → correct_horizontal).
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


def _top(series, n):
    return " || ".join(pd.Series([x for x in series if x and x != "nan"]).drop_duplicates().head(n).tolist())


def dump(ent, year, unclassified, inputs_only, batch, top):
    com = ENTITIES[ent]
    p = ROOT / "data" / ent / "interim" / f"{com}_tagged_rows.parquet"
    if not p.exists():
        print(f"[{ent}] X missing {p.relative_to(ROOT)} — run kedro first.")
        return
    cf = yaml.safe_load((ROOT / "conf" / com / "parameters.yml").open(encoding="utf-8"))
    cols = cf.get("columns", {}) or {}
    df = pq.read_table(p).replace_schema_metadata(None).to_pandas()
    if not {"horizontal_id", "report_period"} <= set(df.columns):
        print(f"[{ent}] X tagged_rows missing horizontal_id/report_period")
        return

    ac, ad = fuzzy(df, cols.get("account_code")), fuzzy(df, cols.get("account_desc"))
    de, am, pj = fuzzy(df, cols.get("description")), fuzzy(df, cols.get("amount")), fuzzy(df, cols.get("project"))
    # step4 DROPS the signature column from tagged_rows (RAM reclaim) — rebuild it identically
    # via the same shared helper (respects conf signature_fields; default acct|adesc|desc_norm[|job]).
    sys.path.insert(0, str(ROOT / "src"))
    from kpi.lib.text import normalize_description, resolve_signature_fields, compose_signature
    jc = fuzzy(df, cols.get("job_code"))
    _empty = pd.Series([""] * len(df), index=df.index)
    _parts = {
        "account_code": (df[ac].astype("string").fillna("").str.strip() if ac else _empty),
        "account_desc": (df[ad].astype("string").fillna("").str.strip() if ad else _empty),
        "desc_norm": (df[de].apply(normalize_description) if de else _empty),
        "job_code": (df[jc].astype("string").fillna("").str.strip() if jc else _empty),
    }
    df["signature"] = compose_signature(_parts, resolve_signature_fields(cf, cols, df.columns))
    df["_H"] = df["horizontal_id"].astype(str)
    df["_ad"] = df[ad].astype(str) if ad else ""
    df["_amt"] = pd.to_numeric(df[am], errors="coerce").fillna(0) if am else 0.0

    # peer_H across ALL years: most-common non-H_OTHER H per account_desc
    cl = df[~df["_H"].isin(["H_OTHER", "", "nan", "None"])]
    peer = cl.groupby("_ad")["_H"].agg(lambda s: s.mode().iloc[0] if len(s.mode()) else "")

    yr = df["report_period"].astype(str)
    sub = df[yr.str.startswith(str(year))].copy()
    if sub.empty:
        print(f"[{ent}] no rows for year {year}. report_period seen: {sorted(yr.unique())[:10]}")
        return
    sub["_ac"] = sub[ac].astype(str) if ac else ""
    sub["_de"] = sub[de].astype(str) if de else ""
    sub["_pj"] = sub[pj].astype(str) if pj else ""

    g = sub.groupby("signature")
    out = g.agg(
        current_H=("_H", lambda s: s.mode().iloc[0] if len(s.mode()) else "H_OTHER"),
        account_code=("_ac", "first"),
        account_desc=("_ad", "first"),
        n_rows=("_H", "size"),
        amount=("_amt", "sum"),
    ).reset_index()
    out["desc_sample"] = out["signature"].map(g["_de"].agg(lambda s: _top(s, 3))).str.slice(0, 90)
    out["projects"] = out["signature"].map(g["_pj"].agg(lambda s: _top(s, 2))).str.slice(0, 50)
    out["peer_H"] = out["account_desc"].map(peer).fillna("")
    if inputs_only:
        # BLIND classification input — hide current_H/peer_H so the LLM pass is independent
        # of the existing (項目組-similar) answer. Compare happens AFTER, in a separate join.
        out = out[["signature", "account_code", "account_desc", "desc_sample", "n_rows", "amount", "projects"]]
    else:
        out = out[["signature", "current_H", "peer_H", "account_code", "account_desc",
                   "desc_sample", "n_rows", "amount", "projects"]]

    tot = len(out)
    n_other = int(out["current_H"].eq("H_OTHER").sum())
    amt_other = float(out.loc[out["current_H"].eq("H_OTHER"), "amount"].abs().sum())
    amt_all = float(out["amount"].abs().sum())
    if unclassified:
        out = out[out["current_H"].eq("H_OTHER")].copy()
    out = out.reindex(out["amount"].abs().sort_values(ascending=False).index)

    out_dir = ROOT / "data" / "review" / "_dump"
    out_dir.mkdir(parents=True, exist_ok=True)
    suffix = f"_{year}_unclassified" if unclassified else f"_{year}"
    tsv = out_dir / f"{ent}_sigs{suffix}.tsv"
    out.to_csv(tsv, sep="\t", index=False, encoding="utf-8-sig")
    n = len(out)
    n_batches = (n + batch - 1) // batch if n else 0
    for b in range(n_batches):
        out.iloc[b * batch:(b + 1) * batch].to_csv(
            out_dir / f"{ent}_sigs{suffix}_batch{b + 1:02d}of{n_batches:02d}.tsv",
            sep="\t", index=False, encoding="utf-8-sig")

    print(f"\n===== {ent} year {year}: {tot:,} sigs total | H_OTHER {n_other:,} "
          f"(|Σamt| {amt_other:,.0f} = {amt_other / amt_all * 100 if amt_all else 0:.0f}% of year) =====")
    print(f"  dumped {n:,} rows{' (H_OTHER only)' if unclassified else ''} → {tsv.relative_to(ROOT)}")
    print(f"  {n_batches} paste-batch files (≤{batch}): {ent}_sigs{suffix}_batch01of{n_batches:02d}.tsv ...")
    print(f"----- top {min(top, n)} by |amount| -----")
    print(out.head(top).to_csv(sep="\t", index=False))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", nargs="+", required=True, choices=sorted(ENTITIES))
    ap.add_argument("--year", required=True, help="report year prefix, e.g. 23")
    ap.add_argument("--unclassified", action="store_true", help="only current H == H_OTHER (the gap to fill)")
    ap.add_argument("--inputs-only", action="store_true",
                    help="hide current_H/peer_H → blind input for independent LLM classification")
    ap.add_argument("--batch", type=int, default=3000)
    ap.add_argument("--top", type=int, default=40)
    a = ap.parse_args()
    for e in a.entity:
        dump(e, a.year, a.unclassified, a.inputs_only, a.batch, a.top)


if __name__ == "__main__":
    main()
