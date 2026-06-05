"""inspect_vproj.py — per-PROJECT vertical (V) review: OUR vertical_label vs the project NAME
(+ 項目組's own label 類別2 if present). NG is shown for context but is FIXED (from 投資領域 /
databook — never override NG). Obvious mis-V'd projects pop out → row_vertical_overrides (V only).

Works for ANY entity/year. If the 項目組 V column (類別2) is absent (every entity except VML 23),
it just shows  project | NG | our_V | Σamt  so we eyeball our_V against the project name.

  python scripts/inspect_vproj.py --entity vml --year 23            # has 類別2 → 4-col compare
  python scripts/inspect_vproj.py --entity galaxy --year 25         # no 類別2 → eyeball our_V vs name
  python scripts/inspect_vproj.py --entity sjm --year 24 --top 80
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


def _cn_kw(s) -> str:
    s = str(s)
    if "非博彩" in s: return ""
    for kws, ng in [(["博彩", "gaming"], "NG0"), (["海上"], "NG10"), (["外國", "客源", "國際客"], "NG1"),
                    (["會議", "會展", "mice"], "NG2"), (["娛樂", "演唱", "表演"], "NG3"), (["體育", "賽事"], "NG4"),
                    (["文化", "藝術", "文藝"], "NG5"), (["健康", "養生"], "NG6"), (["主題", "遊樂"], "NG7"),
                    (["美食", "餐飲"], "NG8"), (["社區"], "NG9"), (["其他"], "NG11")]:
        if any(k in s or k in s.lower() for k in kws): return ng
    return ""


def ng_cols_of(df, cf, cols):
    names = [cols.get("ng11_category", "")] + [
        (ys.get("columns_override") or {}).get("ng11_category") for ys in (cf.get("yearly_sources") or [])]
    out = []
    for nm in names:
        fc = find(df, exact=nm)
        if fc and fc not in out: out.append(fc)
    for c in df.columns:
        if c not in out and any(k in str(c) for k in ("項目性質", "項目類型", "項目分類", "範疇", "投資領域", "NG11 Category")):
            out.append(c)
    return out


def derive_ng(df, ngcols):
    try:
        sys.path.insert(0, str(ROOT / "src"))
        from kpi.lib.conf import load_categories
        from kpi.pipelines.step2_tag_projects._logic import normalize_ng_code
        cats = load_categories()
        def res(x):
            for c in (x, x.upper().replace(" ", "")):
                r = normalize_ng_code(c, cats) or ""
                if r[:2] == "NG" and r[2:].isdigit(): return r
            return _cn_kw(x)
    except Exception:
        res = _cn_kw
    out = pd.Series("", index=df.index, dtype="object")
    for fc in ngcols:
        m = {x: res(x) for x in set(df[fc].astype(str).unique())}
        r = df[fc].astype(str).map(m).fillna("")
        r = r.where(r.str.fullmatch(r"NG\d+").fillna(False), "")
        out = out.mask(out.eq(""), r)
    return out.replace("", "(未分類)")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", required=True, choices=sorted(ENT))
    ap.add_argument("--year", required=True)
    ap.add_argument("--theircol", default=None, help="項目組 V column (default auto: 類別2)")
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
    if df.empty:
        print(f"X no rows for {a.entity} {a.year}"); return

    amt = next((c for c in [cols.get("amount"), "MOP Amt", "調整後金額", "Reported Amount(MOP)"]
                if c and c in df.columns and numify(df[c]).abs().sum() > 0), None)
    df["_amt"] = numify(df[amt]) if amt else 0.0
    proj = find(df, exact=cols.get("project")) or find(df, "SubProject_Name", "Project", "項目")
    ourv = "vertical_label" if "vertical_label" in df.columns else "vertical_id"
    their = a.theircol or find(df, exact="類別2") or find(df, "類別2", "类别2", "類別1", "类别1")
    df["_ng"] = derive_ng(df, ng_cols_of(df, cf, cols))
    print(f"[{a.entity} {a.year}] project={proj!r}  our_V={ourv!r}  項目組_V={their!r}  amount={amt!r}")
    if not proj:
        print("  ⚠ no project column found"); return

    df["_p"] = df[proj].astype(str).str.strip()
    rows = []
    for p, g in df.groupby("_p"):
        rows.append({"project": p, "ng": mode(g["_ng"]), "our_V": mode(g[ourv]),
                     "their_V": mode(g[their]) if their else "", "amt": g["_amt"].sum()})
    out = pd.DataFrame(rows)
    out = out.reindex(out["amt"].abs().sort_values(ascending=False).index)
    n = len(out) if a.top == "all" else int(a.top)
    has_their = bool(their) and out["their_V"].str.strip().ne("").any()
    print(f"\n=== per-project V review  ({len(out)} projects, showing {n})  — NG is FIXED, eyeball our_V ===")
    if has_their:
        print("  project | NG | our_V | 項目組_類別2 | Σamt")
        for _, r in out.head(n).iterrows():
            print(f"  {str(r['project'])[:32]:34s} {r['ng']:<7} {str(r['our_V'])[:13]:15s} "
                  f"{str(r['their_V'])[:16]:18s} {r['amt']:>15,.0f}")
    else:
        print("  project | NG | our_V | Σamt   (no 項目組 V col — eyeball our_V vs the project name)")
        for _, r in out.head(n).iterrows():
            print(f"  {str(r['project'])[:40]:42s} {r['ng']:<7} {str(r['our_V'])[:16]:18s} {r['amt']:>15,.0f}")


if __name__ == "__main__":
    main()
