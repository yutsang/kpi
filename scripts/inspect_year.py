"""inspect_year.py — diagnose a year's numbers for an entity (why 未分類 / double-count / unique-count).

  # Galaxy 24: where is the 未分類 (372M) coming from? show raw NG-source values for blank rows
  python scripts/inspect_year.py --entity galaxy --year 24
  # VML 24: is the 4.6B a double-count? biggest sigs + duplicate unique_id (ref) radar
  python scripts/inspect_year.py --entity vml --year 24
  # audit-master 次數/unique count diagnosis: which col feeds it, is it populated, count per NG×V
  python scripts/inspect_year.py --entity vml --year 24 --unique
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


def _cn_kw(s) -> str:
    s = str(s)
    if "非博彩" in s: return ""
    for kws, ng in [(["博彩", "gaming"], "NG0"), (["海上"], "NG10"), (["外國", "客源", "國際客"], "NG1"),
                    (["會議", "會展", "mice"], "NG2"), (["娛樂", "演唱", "表演"], "NG3"), (["體育", "賽事"], "NG4"),
                    (["文化", "藝術", "文藝"], "NG5"), (["健康", "養生"], "NG6"), (["主題", "遊樂"], "NG7"),
                    (["美食", "餐飲"], "NG8"), (["社區"], "NG9"), (["其他"], "NG11")]:
        if any(k in s or k in s.lower() for k in kws): return ng
    return ""


def fuzzy(df, name):
    if not name: return None
    if name in df.columns: return name
    for c in df.columns:
        if str(c).strip() == str(name).strip(): return c
    return None


def numify(s):
    return pd.to_numeric(s.astype(str).str.replace(",", "", regex=False).str.replace(r"^\s*-\s*$", "0", regex=True),
                         errors="coerce").fillna(0)


def best_amount(df, cands):
    """first candidate (priority order) with non-zero sum — prefers pipeline cols['amount']."""
    for c in cands:
        if c and c in df.columns and numify(df[c]).abs().sum() > 0:
            return c
    return None


def ng_cols_of(df, cf, cols):
    names = [cols.get("ng11_category", "")] + [
        (ys.get("columns_override") or {}).get("ng11_category") for ys in (cf.get("yearly_sources") or [])]
    out = []
    for nm in names:
        fc = fuzzy(df, nm)
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
    p = argparse.ArgumentParser()
    p.add_argument("--entity", required=True, choices=sorted(ENT))
    p.add_argument("--year", required=True)
    p.add_argument("--unique", action="store_true", help="diagnose the audit-master 次數/unique count")
    a = p.parse_args()
    com = ENT[a.entity]
    # tagged_rows has ALL raw cols (Subproject, 分類1, etc.) — kpi_report curates them out
    src = ROOT / "data" / a.entity / "interim" / f"{com}_tagged_rows.parquet"
    if not src.exists():
        src = ROOT / "data" / a.entity / "output" / f"{com}_kpi_report.parquet"
    if not src.exists():
        print(f"X missing tagged_rows/kpi_report for {a.entity}"); return
    cf = yaml.safe_load((ROOT / "conf" / com / "parameters.yml").open(encoding="utf-8"))
    cols = cf.get("columns", {}) or {}
    df = pq.read_table(src).replace_schema_metadata(None).to_pandas()
    amt = best_amount(df, [cols.get("amount"), "MOP Amt", "調整後金額", "Reported Amount(MOP)",
                           fuzzy(df, cols.get("amount")), next((c for c in df.columns if re.search(r"amount|金額", str(c), re.I)), None)])
    df["_amt"] = numify(df[amt]) if amt else 0.0
    print(f"  (source={src.name}, amount_col={amt!r})")
    ngcols = ng_cols_of(df, cf, cols)
    df["_ng"] = derive_ng(df, ngcols)
    yc = next((c for c in ("report_period", "report_year") if c in df.columns), None)
    yr = df[yc].astype(str) if yc else pd.Series("", index=df.index)
    sub = df[yr.str.startswith(a.year)].copy()
    tot = float(sub["_amt"].sum())
    print(f"\n[{a.entity} {a.year}] rows={len(sub):,}  Σ={tot:,.0f}   NG source col(s)={ngcols}")
    ac = fuzzy(df, cols.get("account_code")); ad = fuzzy(df, cols.get("account_desc"))
    pj = fuzzy(df, cols.get("project")); ui = fuzzy(df, cols.get("unique_id"))

    if a.unique:
        print("\n===== UNIQUE / 次數 DIAGNOSIS (audit-master counts distinct subproject per NG×V) =====")
        _cc = cols.get("count_cols")
        _cc = [_cc] if isinstance(_cc, str) else (_cc or [])
        cand = [cols.get("project", "")] + (cols.get("project_name_cols") or []) + _cc
        for nm in [c for c in dict.fromkeys(cand) if c]:
            fc = fuzzy(df, nm)
            if not fc:
                print(f"  conf col {nm!r}: NOT FOUND in parquet ⚠"); continue
            s = sub[fc].astype(str).str.strip()
            blank = (s.eq("") | s.eq("nan") | s.eq("None")).mean() * 100
            print(f"  col {nm!r:24s} -> distinct={s.nunique():>6,}  blank={blank:4.0f}%{'  ⚠ALL/MOST BLANK' if blank>50 else ''}")
        vl = "vertical_label" if "vertical_label" in sub.columns else "vertical_id"
        subc = fuzzy(df, (cols.get("project_name_cols") or [None])[0]) or pj
        if subc:
            sub["_sp"] = sub[subc].astype(str).str.strip()
            g = sub.groupby([ "_ng", vl])["_sp"].nunique().reset_index(name="次數")
            g = g[g["次數"] > 0].sort_values("次數", ascending=False)
            print(f"\n  次數 (distinct {subc!r}) per NG×V — top 20:")
            for _, r in g.head(20).iterrows():
                print(f"     {r['_ng']:<8} {str(r[vl])[:18]:20s} 次數={int(r['次數'])}")
            print(f"  total (NG×V) rows with 次數>0: {len(g)} | 次數 all=1? {bool((g['次數']==1).all())}")
        return

    # NG breakdown + 未分類 drill
    ng = sub.groupby("_ng")["_amt"].agg(["size", "sum"])
    ng = ng.reindex(ng["sum"].abs().sort_values(ascending=False).index)
    print("\n=== NG breakdown ===")
    for k, r in ng.iterrows():
        print(f"  {k:<10} {r['sum']:>16,.0f}  {r['sum']/tot*100 if tot else 0:5.1f}%  ({int(r['size']):,} rows)")
    unc = sub[sub["_ng"].eq("(未分類)")]
    if len(unc):
        print(f"\n=== 未分類 drill: {len(unc):,} rows, Σ={unc['_amt'].sum():,.0f} ===")
        print("  raw NG-source values on 未分類 rows (why blank):")
        for fc in ngcols:
            vc = unc[fc].astype(str).str.strip().replace("", "(empty)").value_counts().head(8)
            print(f"    col {fc!r}: " + ", ".join(f"{v!r}×{n}" for v, n in vc.items()))
        if ac:
            print("  top 未分類 by account_code:")
            for v, s in unc.groupby(unc[ac].astype(str))["_amt"].sum().abs().sort_values(ascending=False).head(10).items():
                print(f"     {str(v)[:26]:28s} Σ={sub[sub[ac].astype(str).eq(v)]['_amt'].sum():,.0f}")

    # double-count radar
    print("\n=== double-count radar ===")
    sig = (sub[ad].astype(str) if ad else "") + " / " + (sub[pj].astype(str).str.slice(0,24) if pj else "")
    big = sub.assign(_s=sig).groupby("_s")["_amt"].sum()
    big = big.reindex(big.abs().sort_values(ascending=False).index).head(8)
    for s, v in big.items():
        print(f"  {str(s)[:54]:56s} {v:>15,.0f}  {v/tot*100 if tot else 0:4.1f}%")
    if ui:
        u = sub[ui].astype(str)
        dup = sub[u.duplicated(keep=False) & u.ne("") & u.ne("nan")]
        if len(dup):
            net = dup["_amt"].sum()
            print(f"\n  duplicate unique_id ({ui}): {len(dup):,} rows share a ref | their Σ={net:,.0f}")
            print(f"   (Σ≈0 → offsetting adjustment pairs, NOT double-count; Σ large → possible double-count)")


if __name__ == "__main__":
    main()
