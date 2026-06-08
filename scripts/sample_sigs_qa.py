"""Signature-level ACCURACY QA — the real bar (not 'every row has a label'). For each entity it
groups the tagged rows by (account_code | account_desc) and dumps the TOP-N groups by |amount|
with their currently-assigned Horizontal/Vertical, so each big-$ signature can be eyeballed for
CORRECTNESS (is this account/desc really → this H?). Sorted by Σ|amount| (the $ that matters most).

Columns map is read from each conf/company_N/parameters.yml (account_code / account_desc /
description / ng11_category / amount). horizontal_label / vertical_label come from tagged_rows.

Run (Windows):
  python scripts/sample_sigs_qa.py --year 23 --top 150           # all 6 entities, 2023
  python scripts/sample_sigs_qa.py --entity galaxy --year 25     # one entity/year
Output: results/sample_sigs_qa_<ent>_<year>.tsv  (+ printed summary)
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd, yaml

ROOT = Path(__file__).resolve().parent.parent
ENTS = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
        "vml": "company_4", "melco": "company_5", "mgm": "company_6"}


def confcols(comp):
    p = ROOT / "conf" / comp.replace("company_", "company_") / "parameters.yml"
    p = ROOT / "conf" / comp / "parameters.yml"
    c = (yaml.safe_load(p.read_text(encoding="utf-8")) or {}).get("columns", {}) if p.exists() else {}
    return c


def pick(df, *names):
    for n in names:
        if n and n in df.columns: return n
    norm = {str(x).strip(): x for x in df.columns}
    for n in names:
        if n and str(n).strip() in norm: return norm[str(n).strip()]
    return None


def modal(s):
    s = s[s.astype(str).str.strip().ne("")]
    return s.value_counts().index[0] if len(s) else ""


def dump_ent(alias, comp, year, top, L):
    tr = ROOT / "data" / alias / "interim" / f"{comp}_tagged_rows.parquet"
    if not tr.exists():
        L.append(f"## {alias}: X {tr} missing"); return None
    df = pd.read_parquet(tr)
    cc = confcols(comp)
    amt = pick(df, cc.get("amount"), "amount_mop")
    ac = pick(df, cc.get("account_code")); ad = pick(df, cc.get("account_desc"))
    desc = pick(df, cc.get("description")); ng = pick(df, cc.get("ng11_category"))
    proj = pick(df, cc.get("project"))
    per = pick(df, "report_period", "report_year", "years")
    if per and year:
        df = df[df[per].astype(str).str.startswith(str(year))].copy()
    if not len(df):
        L.append(f"## {alias} {year}: 0 rows"); return None
    a = pd.to_numeric(df[amt], errors="coerce").fillna(0.0) if amt else pd.Series(0.0, index=df.index)
    g = pd.DataFrame({
        "_ac": df[ac].astype("string").fillna("").str.strip() if ac else "",
        "_ad": df[ad].astype("string").fillna("").str.strip() if ad else "",
        "_desc": df[desc].astype("string").fillna("").str.strip() if desc else "",
        "_proj": df[proj].astype("string").fillna("").str.strip() if proj else "",
        "_ng": df[ng].astype("string").fillna("").str.strip() if ng else "",
        "_h": df["horizontal_label"].astype("string").fillna("") if "horizontal_label" in df.columns else "",
        "_v": df["vertical_label"].astype("string").fillna("") if "vertical_label" in df.columns else "",
        "_a": a.values, "_abs": a.abs().values,
    })
    tot = g["_abs"].sum() or 1
    grp = g.groupby(["_ac", "_ad"])
    rows = []
    for (kac, kad), sub in grp:
        rows.append({
            "Σamt": round(sub["_a"].sum()), "abs": sub["_abs"].sum(), "rows": len(sub),
            "account_code": kac, "account_desc": kad,
            "H_now": modal(sub["_h"]), "V_now": modal(sub["_v"]), "NG": modal(sub["_ng"]),
            "desc_eg": " | ".join(map(str, pd.Series(sub["_desc"][sub["_desc"].ne("")].unique())[:2]))[:80],
            "proj_eg": " | ".join(map(str, pd.Series(sub["_proj"][sub["_proj"].ne("")].unique())[:2]))[:50],
        })
    out = pd.DataFrame(rows).sort_values("abs", ascending=False).head(top)
    out["%yr"] = (out["abs"] / tot * 100).round(2)
    out["audit_H?"] = ""   # ← I fill: correct / →H_XXX
    cols = ["%yr", "Σamt", "rows", "account_code", "account_desc", "NG", "H_now", "V_now", "desc_eg", "proj_eg", "audit_H?"]
    fp = ROOT / "results" / f"sample_sigs_qa_{alias}_{year}.tsv"
    fp.parent.mkdir(exist_ok=True)
    out[cols].to_csv(fp, sep="\t", index=False, encoding="utf-8-sig")
    covered = out["abs"].sum() / tot * 100
    n_other = int((out["H_now"].astype(str).str.contains("其他|OTHER", na=False)).sum())
    L.append(f"## {alias} {year}: top {len(out)} sigs cover {covered:.1f}% of |amt|  "
             f"(H_OTHER-ish in top: {n_other})  → {fp.name}")
    return fp


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", default=None)
    ap.add_argument("--year", default="23")
    ap.add_argument("--top", type=int, default=150)
    a = ap.parse_args()
    L = [f"# sample_sigs_qa  year={a.year}  top={a.top}  (eyeball H_now correctness per big-$ signature)"]
    items = [(a.entity, ENTS[a.entity])] if a.entity else list(ENTS.items())
    for alias, comp in items:
        try: dump_ent(alias, comp, a.year, a.top, L)
        except Exception as e: L.append(f"## {alias}: X {e}")
    print("\n".join(L))
    print(f"\nwrote results/sample_sigs_qa_*_{a.year}.tsv — paste the per-entity TSVs back for audit")


if __name__ == "__main__":
    main()
