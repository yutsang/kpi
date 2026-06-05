"""inspect_vml_golden.py — find the project-team's GOLDEN / 估算 column(s) in the VML data and
reconcile OUR totals against them, per year + per NG.

The golden is usually a PER-PROJECT estimate (one number repeated on every row of a project), so we
report each numeric column two ways: raw Σ (sum of all rows) and per-project-deduped Σ (one value
per project) — the deduped figure is what ties to a project-team target. Columns whose deduped Σ is
within 25% of OUR Σ for a year are flagged ◀ as likely golden.

  python scripts/inspect_vml_golden.py                 # vml, all years
  python scripts/inspect_vml_golden.py --year 24
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
ENT = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
       "vml": "company_4", "melco": "company_5", "mgm": "company_6"}
GOLD_RE = re.compile(r"估算|預算|目標|計劃|計划|核準|核准|批准|承諾|金額|target|plan|budget|golden|approv|commit|投資額|總額", re.I)


def numify(s):
    return pd.to_numeric(s.astype(str).str.replace(",", "", regex=False).str.replace(r"^\s*-\s*$", "0", regex=True),
                         errors="coerce")


def find(df, *subs, exact=None):
    if exact and exact in df.columns:
        return exact
    for c in df.columns:
        if any(s in str(c) for s in subs):
            return c
    return None


def _cn_kw(s) -> str:
    s = str(s)
    if "非博彩" in s: return ""
    for kws, ng in [(["博彩", "gaming"], "NG0"), (["海上"], "NG10"), (["外國", "客源", "國際客"], "NG1"),
                    (["會議", "會展", "mice"], "NG2"), (["娛樂", "演唱", "表演"], "NG3"), (["體育", "賽事"], "NG4"),
                    (["文化", "藝術", "文藝"], "NG5"), (["健康", "養生"], "NG6"), (["主題", "遊樂"], "NG7"),
                    (["美食", "餐飲"], "NG8"), (["社區"], "NG9"), (["其他"], "NG11")]:
        if any(k in s or k in s.lower() for k in kws): return ng
    return ""


def derive_ng(df, cf, cols):
    names = [cols.get("ng11_category", "")] + [
        (ys.get("columns_override") or {}).get("ng11_category") for ys in (cf.get("yearly_sources") or [])]
    ngcols = []
    for nm in names:
        fc = find(df, exact=nm)
        if fc and fc not in ngcols: ngcols.append(fc)
    for c in df.columns:
        if c not in ngcols and any(k in str(c) for k in ("項目性質", "投資領域", "範疇", "NG11 Category")):
            ngcols.append(c)
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
        out = out.mask(out.eq(""), r.where(r.str.fullmatch(r"NG\d+").fillna(False), ""))
    return out.replace("", "(未分類)")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", default="vml", choices=sorted(ENT))
    ap.add_argument("--year", default=None)
    a = ap.parse_args()
    com = ENT[a.entity]
    src = ROOT / "data" / a.entity / "interim" / f"{com}_tagged_rows.parquet"
    if not src.exists():
        print(f"X missing {src.relative_to(ROOT)} — run kedro."); return
    cf = yaml.safe_load((ROOT / "conf" / com / "parameters.yml").open(encoding="utf-8"))
    cols = cf.get("columns", {}) or {}
    df = pq.read_table(src).replace_schema_metadata(None).to_pandas()
    amt = next((c for c in [cols.get("amount"), "MOP Amt", "調整後金額"] if c and c in df.columns
                and numify(df[c]).abs().sum() > 0), None)
    df["_amt"] = numify(df[amt]).fillna(0) if amt else 0.0
    df["_ng"] = derive_ng(df, cf, cols)
    proj = find(df, exact=cols.get("project")) or find(df, "SubProject_Name", "Project")
    yc = next((c for c in ("report_period", "report_year") if c in df.columns), None)
    yr = df[yc].astype(str).str[:2] if yc else pd.Series("", index=df.index)
    years = [a.year] if a.year else sorted(set(yr[yr.str.fullmatch(r"\d{2}")]))

    print(f"[{a.entity}] tagged_rows  amount={amt!r}  project={proj!r}  year_col={yc!r}")
    # ---- candidate golden columns ----
    cands = []
    for c in df.columns:
        if c in ("_amt", "_ng") or c == amt:
            continue
        s = numify(df[c])
        if s.notna().mean() > 0.30 and s.abs().sum() > 0 and (GOLD_RE.search(str(c)) or s.abs().sum() > 1e7):
            cands.append(c)
    print(f"\ncandidate numeric columns ({len(cands)}): " + ", ".join(repr(str(c)) for c in cands[:30]))

    for y in years:
        sub = df[yr.eq(y)]
        if sub.empty: continue
        our = sub["_amt"].sum()
        print(f"\n================  {a.entity} {y}  ================")
        print(f"  OUR Σ (amount={amt!r}) = {our:,.0f}   rows={len(sub):,}")
        # per-NG
        ng = sub.groupby("_ng")["_amt"].sum()
        ng = ng.reindex(ng.abs().sort_values(ascending=False).index)
        print("  OUR per-NG: " + "  ".join(f"{k}={v/1e6:,.0f}M" for k, v in ng.items() if abs(v) > 5e5))
        # each candidate: raw Σ + per-project-deduped Σ
        print("  --- candidate golden columns (raw Σ | per-project-deduped Σ | ◀ if ≈ our) ---")
        rows = []
        for c in cands:
            s = numify(sub[c])
            raw = s.abs().sum()
            if raw == 0: continue
            ded = 0.0
            if proj:
                tmp = pd.DataFrame({"p": sub[proj].astype(str).str.strip(), "v": s})
                ded = tmp.dropna(subset=["v"]).groupby("p")["v"].first().sum()
            flag = " ◀ likely golden" if our and 0.75 <= abs(ded / our) <= 1.25 else (
                   " ◀ raw≈our" if our and 0.75 <= abs(s.sum() / our) <= 1.25 else "")
            rows.append((abs(ded) if ded else raw, c, s.sum(), ded, flag))
        for _, c, raw, ded, flag in sorted(rows, key=lambda x: -x[0])[:12]:
            print(f"     {str(c)[:34]:36s} raw={raw:>16,.0f}  dedup={ded:>16,.0f}{flag}")


if __name__ == "__main__":
    main()
