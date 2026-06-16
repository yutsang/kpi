r"""Name-fill the DICJ code where raw didn't carry it, then SHOW THE EFFECT by tying amounts to golden.

Resolve order per our row (subproject grain):
  1. existing dicj_code (raw 已有)
  2. golden DICJ code is a PREFIX of our project name  (wynn: 'OPCG006_…' → OPCG006)
  3. golden 項目名稱 == our project name (normalized exact)
  4. golden 項目名稱 ⊂ our project name (normalized substring)

Then per entity, BEFORE (existing dicj only) vs AFTER (name-filled): amount coverage % and project-level
tie vs golden 'Database combine' (報告 for 25-buckets / 調整後 for 24·23). amount = conf columns.amount ÷1e4.

Output: prints + results/build_dicj_fill.txt + per-entity dicj_lookup_<ent>.tsv (subproject→DICJ,how)
+ dicj_unmatched_<ent>.tsv (still-blank subprojects + Σamt → 手補). Golden read-only, never committed.

Run (Windows):  python scripts/build_dicj_fill.py            # all
                python scripts/build_dicj_fill.py wynn       # one
"""
from __future__ import annotations
import sys, re, unicodedata
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import yaml

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
PROJ_CANDS = ["Name of Investment Project", "Project Name", "Project name - Amended", "SubProject_Name",
              "Project_code", "項目名稱", "Project", "project", "Subproject"]


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
    for a in "“”‘’\"'（）()「」[]":
        s = s.replace(a, "")
    s = re.sub(r"[\s\-–—_、,，.。/|]+", "", s)
    return s.lower()


def _conf_amt(df, com):
    try:
        a = (yaml.safe_load((ROOT / "conf" / com / "parameters.yml").read_text(encoding="utf-8")).get("columns") or {}).get("amount", "")
        c = _col(df, [a]) if a else None
        if c: return c
    except Exception: pass
    return _col(df, ["Val/COArea Crcy", "MOP Amt", "Amount - Amended", "Debit minus Credit",
                     "Entry Voucher Amount/ Expense Amount", "amount_mop", "amount"])


def main():
    want = [a.lower() for a in sys.argv[1:]] or list(ENT)
    L = ["# build_dicj_fill — 名反填 DICJ + 對 golden 金額睇效果 (before/after)"]
    gpath = next((p for p in GCAND if p.exists()), None)
    if not gpath:
        L.append("!! golden 揾唔到"); _w(L); return
    g = pd.read_excel(gpath, sheet_name=TAB, dtype=str)
    g.columns = [str(c).strip() for c in g.columns]
    g["_a"] = g[_col(g, ["承批公司"])].map(_alias)
    g["_d"] = g[_col(g, ["DICJ Code", "DICJ"])].astype(str).str.strip()
    g["_n"] = g[_col(g, ["項目名稱"])].astype(str).str.strip()
    g["_p"] = g[_col(g, ["Period"])].astype(str).str.strip()
    g["_t"] = g[_col(g, ["Amount Type"])].astype(str).str.strip()
    g["_amt"] = pd.to_numeric(g[_col(g, ["Amount"])].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0.0)
    outdir = ROOT / "results"; outdir.mkdir(exist_ok=True)

    for alias in want:
        com = ENT.get(alias)
        if not com: continue
        L.append(f"\n{'='*70}\n## {alias}")
        ga = g[g["_a"] == alias]
        gname2d = {}
        for _, r in ga[["_d", "_n"]].drop_duplicates().iterrows():
            nn = _norm(r["_n"])
            if nn and nn not in gname2d: gname2d[nn] = r["_d"]
        gnames = sorted(gname2d.keys(), key=len, reverse=True)
        gdset = sorted(set(d for d in ga["_d"].unique() if d and d != "nan"), key=len, reverse=True)
        gdnorm = {_norm(d): d for d in gdset}
        p = ROOT / "data" / alias / "output" / f"{com}_kpi_report.parquet"
        if not p.exists():
            L.append("   (no kpi_report)"); continue
        df = pd.read_parquet(p)
        dcol, pcol, amtc = _col(df, DICJ_CANDS), _col(df, PROJ_CANDS), _conf_amt(df, com)
        ex = df[dcol].astype("string").fillna("").str.strip() if dcol else pd.Series("", index=df.index)
        nm = df[pcol].astype("string").fillna("").str.strip() if pcol else pd.Series("", index=df.index)
        w = pd.to_numeric(df[amtc], errors="coerce").fillna(0.0) / 1e4 if amtc else pd.Series(0.0, index=df.index)
        rp = df["report_period"].astype(str) if "report_period" in df.columns else pd.Series("", index=df.index)

        # resolve per distinct subproject (cache)
        cache = {}
        def resolve(name, exi):
            if exi: return exi, "existing"
            key = name
            if key in cache: return cache[key]
            nn = _norm(name); out = ("", "unmatched")
            for gd in gdset:                                  # 2. golden DICJ code = prefix of name
                if nn.startswith(_norm(gd)):
                    out = (gd, "code_prefix"); break
            else:
                if nn in gname2d: out = (gname2d[nn], "name_exact")
                else:
                    for gn in gnames:
                        if gn and gn in nn: out = (gname2d[gn], "name_substr"); break
            cache[key] = out
            return out

        filled, how = [], []
        for _i in range(len(df)):
            d, h = resolve(nm.iat[_i], ex.iat[_i])
            filled.append(d); how.append(h)
        fd = pd.Series(filled, index=df.index)
        hw = pd.Series(how, index=df.index)

        tot = w.abs().sum()
        cov_ex = w.abs()[ex.ne("")].sum()
        cov_fd = w.abs()[fd.ne("")].sum()
        L.append(f"   amount={amtc!r}  Σ|amt|={tot:,.0f}萬")
        L.append(f"   coverage by amount: existing dicj {cov_ex/max(tot,1)*100:.0f}%  →  name-filled {cov_fd/max(tot,1)*100:.0f}%")
        L.append(f"   resolve how: {hw.value_counts().to_dict()}")

        # tie vs golden (報告 25 / 調整後 24·23), before(existing) vs after(filled)
        gg = ga[ga["_t"].isin(["報告", "調整後"])].copy()
        gg["_use"] = gg["_p"].map(lambda b: "報告" if str(b).startswith("25") else "調整後")
        gsel = gg[gg["_t"] == gg["_use"]]
        gold = gsel.groupby(["_d", "_p"])["_amt"].sum()
        gold.index = gold.index.set_names(["_k", "_b"])
        for tag, key in (("BEFORE(existing)", ex), ("AFTER(name-filled)", fd)):
            o = pd.DataFrame({"_k": key.values, "_b": rp.values, "_w": w.values}).groupby(["_k", "_b"])["_w"].sum()
            j = pd.concat([gold.rename("g"), o.rename("o")], axis=1).fillna(0.0)
            m = j[(j["g"] != 0) & (j["o"] != 0)]
            unmatched = j[(j["g"] != 0) & (j["o"] == 0)]["g"].sum()
            L.append(f"   {tag}: 已配對 Δ={m['o'].sum()-m['g'].sum():,.0f}萬 ({len(m)} key)  "
                     f"golden 對唔到={unmatched:,.0f}萬")

        # diagnose residual (AFTER): bucket vocab + 邊類對唔到 (bucket-label vs 缺 dicj)
        o_fd = pd.DataFrame({"_k": fd.values, "_b": rp.values, "_w": w.values}).groupby(["_k", "_b"])["_w"].sum()
        jj = pd.concat([gold.rename("g"), o_fd.rename("o")], axis=1).fillna(0.0)
        ours_b = sorted(set(str(x) for x in rp.unique() if str(x).strip()))
        gold_b = sorted(set(str(x) for x in gold.index.get_level_values("_b")))
        L.append(f"   our buckets   ={ours_b}")
        L.append(f"   golden buckets={gold_b}")
        our_d = set(x for x in fd.unique() if str(x).strip())
        gold_d = set(gold.index.get_level_values("_k"))
        miss_d = sorted(gold_d - our_d)
        gap = jj[(jj["g"] != 0) & (jj["o"] == 0)].reset_index()
        gap = gap.sort_values("g", key=lambda s: s.abs(), ascending=False)
        gap_in_missd = gap[gap["_k"].isin(miss_d)]["g"].abs().sum()
        gap_buckmis = gap[~gap["_k"].isin(miss_d)]["g"].abs().sum()
        L.append(f"   殘差拆解: 缺dicj(我側完全冇此code)={gap_in_missd:,.0f}萬 ({len(miss_d)} code, e.g.{miss_d[:8]})  "
                 f"|  bucket對唔上(code在但period唔match)={gap_buckmis:,.0f}萬")
        gap.head(200).to_csv(outdir / f"dicj_goldgap_{alias}.tsv", sep="\t", index=False)

        # golden 中文名→DICJ map (供 step0 dicj_name_substr_map_file 反填用, e.g. wynn)
        nmap = ga[["_n", "_d"]].copy()
        nmap = nmap[(nmap["_n"].astype(str).str.strip() != "") & (nmap["_d"].astype(str).str.strip() != "")]
        nmap.drop_duplicates().to_csv(outdir / f"dicj_namemap_{alias}.tsv", sep="\t", index=False, header=False)

        # write lookup + unmatched
        lk = pd.DataFrame({"subproject": nm, "dicj": fd, "how": hw}).drop_duplicates("subproject")
        lk.sort_values(["how", "subproject"]).to_csv(outdir / f"dicj_lookup_{alias}.tsv", sep="\t", index=False)
        un = lk[lk["dicj"] == ""].copy()
        if len(un):
            un["amt萬"] = un["subproject"].map(lambda x: round(w.abs()[nm.eq(x)].sum(), 1))
            un.sort_values("amt萬", ascending=False)[["subproject", "amt萬"]].head(400).to_csv(
                outdir / f"dicj_unmatched_{alias}.tsv", sep="\t", index=False)
    _w(L)


def _w(L):
    out = ROOT / "results" / "build_dicj_fill.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)} + dicj_lookup_*/dicj_unmatched_*  ← paste back")


if __name__ == "__main__":
    main()
