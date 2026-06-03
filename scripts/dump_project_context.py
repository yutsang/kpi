"""Dump per-project step2 context so the classification can be done HERE (Claude) instead of
the internal LLM — then injected as manual_vertical (pins V, no LLM drift).

For each unique project it rebuilds the SAME context step2 feeds its LLM
(`project_context_from_group` + `normalize_ng_code` + `candidates_for_ng`, imported from the
pipeline so they can never drift), and writes one TSV row:

  project | their_性質(raw 項目性質) | NG(normalized) | ng_label | candidates |
  buckets | Σamt | rows | capex_opex | cur_V | top_accounts | top_descs

`their_性質` vs `NG` lets us see which 項目性質 labels FAIL to normalize (NG==raw → unmapped →
that project's NG is unreliable, classify by name/accounts). One row per project = covers ALL
years (V is per-project), so it fixes both 24-untagged AND 25 mis-tags in one pass.

Run (Windows), after kedro has built interim/{company}_raw.parquet + _unique_projects.xlsx:
  python scripts/dump_project_context.py --entity melco
  python scripts/dump_project_context.py --entity wynn
→ results/ctx_{entity}.tsv  (drop into the Mac results/  — or paste).

Claude then classifies each project's V (within `candidates`) and writes
data/_overrides/{entity}_vertical.tsv  →  inject_manual_vertical.py  →  kedro run.
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
ENTITIES = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
            "vml": "company_4", "melco": "company_5", "mgm": "company_6"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", required=True, choices=sorted(ENTITIES))
    ap.add_argument("--topn", type=int, default=6, help="top account/desc samples per project")
    args = ap.parse_args()
    company = ENTITIES[args.entity]

    from kpi.lib.conf import load_categories
    from kpi.pipelines.step2_tag_projects._logic import (
        project_context_from_group, normalize_ng_code, candidates_for_ng)

    cats = load_categories()
    ng_categories = cats.get("ng_categories", {})
    cfg = yaml.safe_load((ROOT / "conf" / company / "parameters.yml").read_text(encoding="utf-8"))
    cols = cfg["columns"]

    interim = ROOT / "data" / args.entity / "interim"
    pq = interim / f"{company}_raw.parquet"
    px = interim / f"{company}_unique_projects.xlsx"
    if not pq.exists() or not px.exists():
        print(f"X missing {pq.name} or {px.name} under data/{args.entity}/interim — "
              f"run kedro (step0+step1) first."); return

    proj_df = pd.read_excel(px)
    pcol = cols["project"]
    if pcol not in proj_df.columns:
        print(f"X project col {pcol!r} not in {px.name}; cols={list(proj_df.columns)}"); return
    projects = proj_df[pcol].astype("string").fillna("").tolist()
    # Extra descriptive name columns (e.g. Wynn project col is a CODE → add 'Sub project'
    # + '项目名称中文' so V is classifiable). project_name_cols from conf; skip the code col itself.
    vcol = next((c for c in ("manual_vertical", "llm_vertical", "vertical_id") if c in proj_df.columns), None)
    cur = dict(zip(proj_df[pcol].astype("string").fillna(""),
                   proj_df[vcol].astype("string").fillna(""))) if vcol else {}

    df = pd.read_parquet(pq)
    # Guard: project_context_from_group reads these 5 cols — create empties if absent so it
    # can't KeyError on an entity whose raw is missing one.
    for key in ("ng11_category", "capex_opex", "account_desc", "description", "amount"):
        c = cols.get(key)
        if c and c not in df.columns:
            print(f"  ! {key} col {c!r} absent in parquet — using blanks")
            df[c] = "" if key != "amount" else 0
    df = df.assign(_proj=df[pcol].astype("string").fillna(""))
    has_bucket = "report_period" in df.columns
    name_cols = [c for c in (cols.get("project_name_cols") or [])
                 if c and c in df.columns and c != pcol]
    if name_cols:
        print(f"  name_hint cols: {name_cols}")
    groups = dict(tuple(df.groupby("_proj", sort=False)))

    def _distinct(series, k=5):
        out = []
        for x in series.tolist():
            s = str(x).strip()
            if s and s.lower() != "nan" and s not in out:
                out.append(s)
            if len(out) >= k:
                break
        return out

    rows = []
    for proj in projects:
        sub = groups.get(proj)
        if sub is None or len(sub) == 0:
            continue
        ctx = project_context_from_group(proj, sub, cols)
        raw_ng = ctx["primary_ng"] or ""
        ng = normalize_ng_code(raw_ng, cats) or "NG11"
        cands = candidates_for_ng(ng, cats)
        co = ctx["capex_opex_counts"]
        buckets = (",".join(sorted(set(sub["report_period"].astype(str)))) if has_bucket else "")
        name_hint = " ¦ ".join(f"{nc}={' / '.join(_distinct(sub[nc]))}"
                               for nc in name_cols if _distinct(sub[nc]))
        rows.append({
            "project": proj,
            "name_hint": name_hint,
            "their_性質": raw_ng,
            "NG": ng,
            "ng_label": ng_categories.get(ng, {}).get("label", ""),
            "candidates": ",".join(cands),
            "buckets": buckets,
            "Σamt": round(ctx["total_amount"]),
            "rows": ctx["row_count"],
            "capex_opex": ";".join(f"{k}:{v}" for k, v in co.items()),
            "cur_V": cur.get(proj, ""),
            "top_accounts": " | ".join(ctx["top_account_descs"][:args.topn]),
            "top_descs": " | ".join(ctx["top_descriptions"][:args.topn]),
        })
    out = pd.DataFrame(rows)
    if out.empty:
        print("X no projects matched between unique_projects.xlsx and raw.parquet"); return
    out = out.reindex(out["Σamt"].abs().sort_values(ascending=False).index).reset_index(drop=True)
    repdir = ROOT / "results"
    repdir.mkdir(parents=True, exist_ok=True)
    rep = repdir / f"ctx_{args.entity}.tsv"
    out.to_csv(rep, sep="\t", index=False, encoding="utf-8-sig")
    # JSONL — ROBUST channel: json escapes tabs/newlines inside strings, so the file
    # survives any tab→space mangling on transfer. THIS is the file to send to Claude.
    # (inject_manual_vertical normalizes whitespace, so internal-space drift still matches.)
    import json as _json
    repj = repdir / f"ctx_{args.entity}.jsonl"
    with repj.open("w", encoding="utf-8") as fh:
        for r in out.to_dict("records"):
            fh.write(_json.dumps(r, ensure_ascii=False) + "\n")
    print(f"[{args.entity}] {len(out)} projects → {rep.relative_to(ROOT)}  +  {repj.relative_to(ROOT)} (send the .jsonl)")

    # NG money histogram
    h = out.groupby("NG")["Σamt"].agg(["size", "sum"]).reindex(
        sorted(out["NG"].unique(), key=lambda x: (len(x), x)))
    print("\nNG histogram (projects / ΣΣamt):")
    print(h.to_string())

    # Which raw 項目性質 values FAIL to normalize (NG == raw means unmapped)?
    unmapped = out[out["NG"] == out["their_性質"]]
    if len(unmapped):
        u = (unmapped.groupby("their_性質")["Σamt"].agg(["size", "sum"])
             .sort_values("sum", key=lambda s: s.abs(), ascending=False))
        print(f"\n⚠ {len(unmapped)} projects whose 項目性質 did NOT map to a NG code "
              f"(→ classify by name/accounts):")
        print(u.head(25).to_string())


if __name__ == "__main__":
    main()
