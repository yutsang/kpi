"""explore.py — drill into the FINAL kpi_report by NG / H / V / account / project / desc.
Made for eyeballing in cmd: filter (AND) → group → top-N (or all). NG is derived from the
databook 項目性質/項目類型/投資領域 column EXACTLY like the deliverable (never from V).

WHEN-GENERATED (answers "when were my files made"):
  python scripts/explore.py --entity vml --files

OVERVIEW:
  python scripts/explore.py --entity vml --year 23 --summary          # NG totals + NG×V + NG×H

DRILL (combine any filters; pick a --by; --top N or 'all'):
  # what projects sit in NG8 (餐飲)?
  python scripts/explore.py --entity vml --year 23 --ng NG8 --by project --top 30
  # ALL accounts tagged H_FNB
  python scripts/explore.py --entity vml --year 23 --h H_FNB --by acctdesc --top all
  # how is one account split across H?
  python scripts/explore.py --entity vml --year 23 --acctdesc "CIP-A&A" --by h
  # what's LEFT as H_OTHER — eyeball the raw rows
  python scripts/explore.py --entity vml --year 23 --h H_OTHER --show examples --top 50
  # a vertical's accounts
  python scripts/explore.py --entity vml --year 23 --v V_CONCERT --by account --top 50

Filters: --year --ng --h --v --account --acctdesc --project --desc   (all substring/prefix, AND-combined)
Views:   --by {row,project,account,acctdesc,ng,h,v,signature}  --show {sum,examples}  --sort {amount,rows}  --top N|all
"""
from __future__ import annotations
import argparse, glob, re, sys, time
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
    if "非博彩" in s:
        return ""
    for kws, ng in [(["博彩", "gaming"], "NG0"), (["海上"], "NG10"),
                    (["外國", "客源", "國際客"], "NG1"), (["會議", "會展", "mice"], "NG2"),
                    (["娛樂", "演唱", "表演"], "NG3"), (["體育", "賽事"], "NG4"),
                    (["文化", "藝術", "文藝"], "NG5"), (["健康", "養生"], "NG6"),
                    (["主題", "遊樂"], "NG7"), (["美食", "餐飲"], "NG8"),
                    (["社區"], "NG9"), (["其他"], "NG11")]:
        if any(k in s or k in s.lower() for k in kws):
            return ng
    return ""


def fuzzy(df, name):
    if not name:
        return None
    if name in df.columns:
        return name
    for c in df.columns:
        if str(c).strip() == str(name).strip():
            return c
    return None


def derive_ng(df, cf, cols):
    names = [cols.get("ng11_category", "")] + [
        (ys.get("columns_override") or {}).get("ng11_category") for ys in (cf.get("yearly_sources") or [])]
    ngcols = []
    for nm in names:
        fc = fuzzy(df, nm)
        if fc and fc not in ngcols:
            ngcols.append(fc)
    for c in df.columns:
        if c not in ngcols and any(k in str(c) for k in ("項目性質", "項目類型", "項目分類", "範疇", "NG11 Category", "NG Category")):
            ngcols.append(c)
    try:
        sys.path.insert(0, str(ROOT / "src"))
        from kpi.lib.conf import load_categories
        from kpi.pipelines.step2_tag_projects._logic import normalize_ng_code
        cats = load_categories()
        def res(x):
            for c in (x, x.upper().replace(" ", "")):
                r = normalize_ng_code(c, cats) or ""
                if r[:2] == "NG" and r[2:].isdigit():
                    return r
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
    p = argparse.ArgumentParser(formatter_class=argparse.RawDescriptionHelpFormatter, description=__doc__)
    p.add_argument("--entity", required=True, choices=sorted(ENT))
    p.add_argument("--files", action="store_true", help="list generated report files + when made")
    p.add_argument("--summary", action="store_true", help="NG totals + NG×V + NG×H")
    p.add_argument("--year"); p.add_argument("--ng"); p.add_argument("--h"); p.add_argument("--v")
    p.add_argument("--account"); p.add_argument("--acctdesc"); p.add_argument("--project"); p.add_argument("--desc")
    p.add_argument("--by", default="row", choices=["row", "project", "account", "acctdesc", "ng", "h", "v", "signature"])
    p.add_argument("--show", default="sum", choices=["sum", "examples"])
    p.add_argument("--sort", default="amount", choices=["amount", "rows"])
    p.add_argument("--top", default="20")
    a = p.parse_args()
    com = ENT[a.entity]

    if a.files:
        print(f"=== generated files for {a.entity} (newest first) ===")
        fs = glob.glob(str(ROOT / "data" / a.entity / "output" / "*")) + \
             glob.glob(str(ROOT / "data" / "review" / f"{a.entity}_*")) + \
             glob.glob(str(ROOT / "data" / "review" / f"*{a.entity}*"))
        for f in sorted(set(fs), key=lambda x: -Path(x).stat().st_mtime):
            st = Path(f).stat()
            print(f"  {time.strftime('%Y-%m-%d %H:%M', time.localtime(st.st_mtime))}  {st.st_size/1e6:6.2f}MB  {Path(f).relative_to(ROOT)}")
        return

    parquet = ROOT / "data" / a.entity / "output" / f"{com}_kpi_report.parquet"
    if not parquet.exists():
        print(f"X missing {parquet.relative_to(ROOT)} — run kedro first."); return
    cf = yaml.safe_load((ROOT / "conf" / com / "parameters.yml").open(encoding="utf-8"))
    cols = cf.get("columns", {}) or {}
    df = pq.read_table(parquet).replace_schema_metadata(None).to_pandas()
    df["_amt"] = pd.to_numeric(df[fuzzy(df, cols.get("amount"))], errors="coerce").fillna(0)
    df["_ng"] = derive_ng(df, cf, cols)
    df["_v"] = df.get("vertical_label", df.get("vertical_id", "")).astype(str)
    df["_h"] = df.get("horizontal_label", df.get("horizontal_id", "")).astype(str)
    df["_hid"] = df.get("horizontal_id", "").astype(str)
    df["_vid"] = df.get("vertical_id", "").astype(str)
    df["_acct"] = df[fuzzy(df, cols.get("account_code"))].astype(str) if fuzzy(df, cols.get("account_code")) else ""
    df["_ad"] = df[fuzzy(df, cols.get("account_desc"))].astype(str) if fuzzy(df, cols.get("account_desc")) else ""
    df["_proj"] = df[fuzzy(df, cols.get("project"))].astype(str) if fuzzy(df, cols.get("project")) else ""
    df["_desc"] = df[fuzzy(df, cols.get("description"))].astype(str) if fuzzy(df, cols.get("description")) else ""
    ycol = next((c for c in ("report_period", "report_year") if c in df.columns), None)

    m = pd.Series(True, index=df.index)
    if a.year and ycol: m &= df[ycol].astype(str).str.startswith(a.year)
    if a.ng:   m &= df["_ng"].str.contains(a.ng, case=False, na=False)
    if a.h:    m &= (df["_hid"].str.contains(a.h, case=False, na=False) | df["_h"].str.contains(a.h, na=False))
    if a.v:    m &= (df["_vid"].str.contains(a.v, case=False, na=False) | df["_v"].str.contains(a.v, na=False))
    if a.account:  m &= df["_acct"].str.contains(a.account, case=False, na=False)
    if a.acctdesc: m &= df["_ad"].str.contains(a.acctdesc, case=False, na=False)
    if a.project:  m &= df["_proj"].str.contains(a.project, case=False, na=False)
    if a.desc:     m &= df["_desc"].str.contains(a.desc, case=False, na=False)
    sub = df[m]
    print(f"[{a.entity}] {len(sub):,} rows match  |  Σamt = {sub['_amt'].sum():,.0f}")
    if sub.empty: return

    if a.summary:
        ng = sub.groupby("_ng")["_amt"].agg(["size", "sum"]).sort_values("sum", ascending=False)
        print("\n=== NG ===");  print(ng.to_string())
        print("\n=== NG × V (top 30) ===")
        print(sub.groupby(["_ng", "_v"])["_amt"].sum().sort_values(ascending=False).head(30).to_string())
        print("\n=== NG × H (top 30) ===")
        print(sub.groupby(["_ng", "_h"])["_amt"].sum().sort_values(ascending=False).head(30).to_string())
        return

    n = len(sub) if a.top == "all" else int(a.top)
    if a.by == "row" or a.show == "examples":
        cols_show = ["_ng", "_v", "_h", "_acct", "_ad", "_proj", "_amt"]
        s = sub.sort_values("_amt", key=lambda x: x.abs(), ascending=False).head(n)
        for _, r in s.iterrows():
            print(f"  {r['_ng']:<9} {r['_v'][:14]:15s} {r['_h'][:12]:13s} {r['_acct'][:14]:15s} "
                  f"{r['_ad'][:24]:26s} {str(r['_proj'])[:30]:32s} {r['_amt']:>14,.0f}")
        return
    key = {"project": "_proj", "account": "_acct", "acctdesc": "_ad", "ng": "_ng", "h": "_h", "v": "_v", "signature": "_ad"}[a.by]
    g = sub.groupby(key)["_amt"].agg(n_rows="size", amount="sum").reset_index()
    g = g.sort_values("amount" if a.sort == "amount" else "n_rows", key=lambda x: x.abs(), ascending=False).head(n)
    print(f"\n=== by {a.by} (top {n if a.top!='all' else 'ALL'}) ===")
    for _, r in g.iterrows():
        print(f"  {str(r[key])[:50]:52s} n_rows={int(r['n_rows']):7,d}  Σ={r['amount']:>15,.0f}")


if __name__ == "__main__":
    main()
