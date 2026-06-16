r"""Fill the project-level DICJ Code onto every row (raw = subproject grain; golden DICJ = project grain).

Sources, in priority order, per our subproject:
  1. existing dicj_code (where the raw already carries it — galaxy 100%, sjm/vml ~23-bucket)
  2. mgm: Project_code = 項目N is itself the DICJ
  3. golden 項目名稱 ↔ our subproject name, NORMALIZED (strip 引號/全半形/空白/–) — exact then substring
The golden 'HQ 投資方向_audit_0616.xlsx' / 'Database combine' gives DICJ ↔ 項目名稱 per 承批公司.

Output (per entity): coverage %, Σamt covered, + results/dicj_lookup_<ent>.tsv (subproject→DICJ, our use)
+ results/dicj_unmatched_<ent>.tsv (subprojects with NO DICJ + Σamt → you fill manually).
Reads only; never writes the pipeline or commits the golden.

Run (Windows):  python scripts/build_dicj_fill.py
"""
from __future__ import annotations
import sys, re, unicodedata
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
GCAND = [ROOT / "HQ 投資方向_audit_0616.xlsx", ROOT / "data" / "HQ 投資方向_audit_0616.xlsx",
         ROOT / "results" / "HQ 投資方向_audit_0616.xlsx"]
TAB = "Database combine"
ENT = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
       "vml": "company_4", "melco": "company_5", "mgm": "company_6"}
NAME2ALIAS = [
    (["galaxy", "銀河"], "galaxy"), (["wynn", "永利"], "wynn"), (["mgm", "美高梅"], "mgm"),
    (["melco", "新濠", "摩珀斯", "影匯", "影滙", "studio city", "city of dreams"], "melco"),
    (["sjm", "澳娛", "葡京", "回力", "上葡京"], "sjm"),
    (["威尼斯", "金沙", "sands", "londoner", "倫敦人", "parisian", "巴黎人", "vml", "venetian"], "vml"),
]
DICJ_CANDS = ["dicj_code", "DICJ Code", "DICJ", "dicj"]
PROJ_CANDS = ["Project Name", "Project", "SubProject_Name", "Project name - Amended", "Project_code",
              "項目名稱", "Name of Investment Project", "project", "Subproject"]
AMT = ["amount_mop", "Val/COArea Crcy", "MOP Amt", "Amount - Amended", "Debit minus Credit",
       "Reported Amount(MOP)", "Entry Voucher Amount/ Expense Amount", "amount"]


def _alias(s):
    sl = str(s).strip().lower()
    for kws, a in NAME2ALIAS:
        if any(k.lower() in sl for k in kws):
            return a
    return None


def _col(df, cands):
    for c in cands:
        if c in df.columns:
            return c
    low = {str(c).strip().lower(): c for c in df.columns}
    for c in cands:
        if str(c).strip().lower() in low:
            return low[str(c).strip().lower()]
    return None


def _norm(s):
    s = unicodedata.normalize("NFKC", str(s)).strip()
    for a, b in [("“", ""), ("”", ""), ("‘", ""), ("’", ""), ('"', ""), ("'", ""),
                 ("（", ""), ("）", ""), ("(", ""), (")", ""), ("「", ""), ("」", "")]:
        s = s.replace(a, b)
    s = re.sub(r"[\s\-–—_、,，.。]+", "", s)
    return s.lower()


def main():
    gpath = next((p for p in GCAND if p.exists()), None)
    if not gpath:
        print("!! golden 揾唔到"); return
    g = pd.read_excel(gpath, sheet_name=TAB, dtype=str)
    g.columns = [str(c).strip() for c in g.columns]
    c_ent = _col(g, ["承批公司"]); c_dicj = _col(g, ["DICJ Code", "DICJ"]); c_proj = _col(g, ["項目名稱"])
    g["_a"] = g[c_ent].map(_alias)
    summary = []
    outdir = ROOT / "results"; outdir.mkdir(exist_ok=True)
    for alias, com in ENT.items():
        ga = g[g["_a"] == alias][[c_dicj, c_proj]].dropna().drop_duplicates()
        ga.columns = ["dicj", "name"]
        # golden lookup: normalized name → DICJ (longest name first for substring safety)
        gmap = {}
        for _, r in ga.iterrows():
            n = _norm(r["name"])
            if n and n not in gmap:
                gmap[n] = str(r["dicj"]).strip()
        gnames = sorted(gmap.keys(), key=len, reverse=True)
        gdicj_norm = {_norm(d): str(d).strip() for d in ga["dicj"] if str(d).strip()}

        p = ROOT / "data" / alias / "output" / f"{com}_kpi_report.parquet"
        if not p.exists():
            print(f"  [{alias}] no kpi_report — skip"); continue
        df = pd.read_parquet(p)
        scol = _col(df, PROJ_CANDS); dcol = _col(df, DICJ_CANDS)
        amtc = _col(df, AMT)
        sub = df[scol].astype("string").fillna("").str.strip() if scol else pd.Series("", index=df.index)
        dex = df[dcol].astype("string").fillna("").str.strip() if dcol else pd.Series("", index=df.index)
        amt = pd.to_numeric(df[amtc], errors="coerce").fillna(0.0).abs() if amtc else pd.Series(0.0, index=df.index)

        # per-subproject: existing dicj (modal) + name-resolve
        tmp = pd.DataFrame({"sub": sub, "dex": dex, "amt": amt})
        # 1. existing dicj per subproject (modal non-blank)
        sub_dicj = {}
        for s, grp in tmp.groupby("sub"):
            if s == "":
                continue
            ex = grp.loc[grp["dex"].ne(""), "dex"]
            sub_dicj[s] = ex.mode().iloc[0] if len(ex) else ""
        # 2. resolve blanks by name (exact → substring → our-name contains golden DICJ code)
        def resolve(s):
            if sub_dicj.get(s):
                return sub_dicj[s], "existing"
            ns = _norm(s)
            if ns in gmap:
                return gmap[ns], "name_exact"
            for gn in gnames:
                if gn and gn in ns:
                    return gmap[gn], "name_substr"
            for dn, dd in gdicj_norm.items():
                if dn and dn in ns:
                    return dd, "code_in_name"
            return "", "unmatched"
        rows = []
        for s in sub_dicj:
            d, how = resolve(s)
            rows.append({"subproject": s, "dicj": d, "how": how})
        lk = pd.DataFrame(rows)
        # coverage by amount
        s2d = dict(zip(lk["subproject"], lk["dicj"]))
        got = sub.map(lambda x: bool(s2d.get(x, "")))
        cov_amt = amt[got].sum() / 1e4
        tot_amt = amt.sum() / 1e4
        n_assigned = int((lk["dicj"] != "").sum())
        how_ct = lk["how"].value_counts().to_dict()
        print(f"[{alias}] subprojects={len(lk):,}  有DICJ={n_assigned:,} ({n_assigned/max(len(lk),1)*100:.0f}%)  "
              f"金額覆蓋={cov_amt:,.0f}/{tot_amt:,.0f}萬 ({cov_amt/max(tot_amt,1)*100:.0f}%)  {how_ct}")
        lk.sort_values(["how", "subproject"]).to_csv(outdir / f"dicj_lookup_{alias}.tsv", sep="\t", index=False)
        # unmatched + amount
        un = lk[lk["dicj"] == ""].copy()
        if len(un):
            un["amt萬"] = un["subproject"].map(lambda x: round(amt[sub.eq(x)].sum() / 1e4, 1))
            un.sort_values("amt萬", ascending=False)[["subproject", "amt萬"]].to_csv(
                outdir / f"dicj_unmatched_{alias}.tsv", sep="\t", index=False)
        summary.append((alias, len(lk), n_assigned, round(cov_amt, 0), round(tot_amt, 0)))

    print("\n# summary  entity | subprojects | 有DICJ | 金額覆蓋萬 | 總額萬")
    for a, n, k, c, t in summary:
        print(f"   {a:7s} {n:>7,} {k:>7,} {c:>12,.0f} {t:>12,.0f}")
    print(f"\nwrote results/dicj_lookup_<ent>.tsv + dicj_unmatched_<ent>.tsv  ← paste unmatched back")


if __name__ == "__main__":
    main()
