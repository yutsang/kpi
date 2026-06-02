"""Stamp the project team's reviewed Vertical onto each entity's unique_projects.xlsx.

The project team reviewed every project and gave us the correct Vertical (the
`manual_vertical` column in data/_overrides/<entity>_vertical.tsv). step4 already
prefers `manual_vertical` over `llm_vertical` (first-match-wins in _final_value),
and step2 SKIPS re-tagging when the project file is already tagged — so injecting
this column and re-running `kedro run --pipeline=<entity>` pins V (hence NG) to the
project team's ground truth, no LLM drift.

Run (Windows), one entity at a time:
  python scripts/inject_manual_vertical.py --entity galaxy
  python scripts/inject_manual_vertical.py --entity galaxy --dry-run   # preview only
Then:
  kedro run --pipeline=galaxy
  python scripts/build_master_audit_25.py --entity galaxy   # (your usual rebuild)

Matching is whitespace/case-insensitive exact on the project name. Any override
project that does not match a row in unique_projects.xlsx is reported as UNMATCHED
so you can see coverage. Nothing is written for unmatched rows.
"""
from __future__ import annotations
import argparse, re, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parent.parent
ENTITIES = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
            "vml": "company_4", "melco": "company_5", "mgm": "company_6"}
TAG_SOURCE = "project_team"   # neutral provenance tag written into the data


def norm(s) -> str:
    return re.sub(r"\s+", " ", str(s).replace('"', " ")).strip().lower()


def project_col_for(company: str) -> str:
    cf = yaml.safe_load((ROOT / "conf" / company / "parameters.yml").open(encoding="utf-8"))
    return (cf.get("columns") or {}).get("project") or "Project"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", required=True, choices=sorted(ENTITIES))
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--clear", action="store_true",
                    help="REVERT: blank manual_vertical (set by a prior inject) so V falls back to llm_vertical")
    args = ap.parse_args()
    company = ENTITIES[args.entity]

    if args.clear:
        alias = args.entity
        interim = ROOT / "data" / alias / "interim"
        hits = sorted(interim.glob("*unique_projects*.xlsx")) if interim.exists() else []
        if not hits:
            hits = [p for p in (ROOT / "data").glob("*/interim/*unique_projects*.xlsx")
                    if alias in str(p).lower() or company in str(p).lower()]
        if not hits:
            print(f"X no *unique_projects*.xlsx for {alias}."); return
        xlsx = hits[0]
        pdf = pd.read_excel(xlsx)
        had = int((pdf.get("manual_vertical", pd.Series(dtype=str)).astype("string").fillna("").str.strip() != "").sum()) if "manual_vertical" in pdf.columns else 0
        if "manual_vertical" in pdf.columns:
            pdf["manual_vertical"] = ""
        if "tag_source" in pdf.columns:
            ts = pdf["tag_source"].astype("string").fillna("")
            pdf["tag_source"] = ts.where(ts != TAG_SOURCE, "")
        pdf.to_excel(xlsx, index=False)
        print(f"[{alias}] cleared manual_vertical on {had} projects in {xlsx.name}. "
              f"Next: kedro run --pipeline={alias}  (V falls back to llm_vertical + conf overrides)")
        return

    ov = ROOT / "data" / "_overrides" / f"{args.entity}_vertical.tsv"
    if not ov.exists():
        print(f"X missing {ov} — copy data/_overrides/ from the Mac side first."); return
    # Robust format: each line is 'V_CODE<ws>project'. V code is one token (no spaces)
    # so it parses whether the tab survived the Mac->Win copy or got turned into spaces.
    omap = {}
    for raw in ov.read_text(encoding="utf-8-sig").splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        parts = line.split(None, 1)            # split on first whitespace run (tab OR space)
        if len(parts) != 2 or not parts[0].startswith("V_"):
            continue
        vcode, proj = parts[0], parts[1]
        omap[norm(proj)] = vcode
    if not omap:
        print(f"X {ov.name} parsed 0 rows — first line: {ov.read_text(encoding='utf-8-sig').splitlines()[:2]!r}")
        return
    print(f"[{args.entity}] override rows: {len(omap)}")

    # Data dir uses the alias (data/galaxy/...); the filename prefix uses the conf
    # code (company_1_unique_projects.xlsx). Glob both so naming quirks don't matter.
    alias = args.entity
    interim = ROOT / "data" / alias / "interim"
    hits = sorted(interim.glob("*unique_projects*.xlsx")) if interim.exists() else []
    if not hits:
        hits = [p for p in (ROOT / "data").glob("*/interim/*unique_projects*.xlsx")
                if alias in str(p).lower() or company in str(p).lower()]
    if not hits:
        print(f"X no *unique_projects*.xlsx under data/{alias}/interim (or data/{company}/interim).")
        print("  run step0+step1+step2 (kedro run) once first."); return
    xlsx = hits[0]
    print(f"  using {xlsx.relative_to(ROOT)}")
    pdf = pd.read_excel(xlsx)
    pcol = project_col_for(company)
    if pcol not in pdf.columns:
        cand = [c for c in pdf.columns if "project" in str(c).lower() or "name" in str(c).lower()]
        print(f"! configured project col {pcol!r} not in xlsx; columns={list(pdf.columns)}")
        if not cand: return
        pcol = cand[0]
        print(f"  falling back to {pcol!r}")

    if "manual_vertical" not in pdf.columns:
        pdf["manual_vertical"] = ""
    if "tag_source" not in pdf.columns:
        pdf["tag_source"] = ""
    pdf["manual_vertical"] = pdf["manual_vertical"].astype("object").fillna("")
    pdf["tag_source"] = pdf["tag_source"].astype("object").fillna("")

    keys = pdf[pcol].map(norm)
    matched = changed = 0
    matched_keys = set()
    for i, k in keys.items():
        v = omap.get(k)
        if not v:
            continue
        matched += 1
        matched_keys.add(k)
        prev = str(pdf.at[i, "manual_vertical"]).strip()
        if prev != v:
            changed += 1
        pdf.at[i, "manual_vertical"] = v
        pdf.at[i, "tag_source"] = TAG_SOURCE

    unmatched = sorted(set(omap) - matched_keys)
    print(f"  projects in xlsx       : {len(pdf)}")
    print(f"  matched & set          : {matched}  (changed value: {changed})")
    print(f"  override keys UNMATCHED : {len(unmatched)}")
    for k in unmatched[:40]:
        print(f"      · {k[:90]}")
    if len(unmatched) > 40:
        print(f"      … +{len(unmatched)-40} more")

    if args.dry_run:
        print("  (dry-run — not written)")
        return
    pdf.to_excel(xlsx, index=False)
    print(f"  wrote manual_vertical into {xlsx.name}.  Next: kedro run --pipeline={args.entity}")


if __name__ == "__main__":
    main()
