r"""Project-level tie vs the NEW golden 'HQ 投資方向_audit_0616.xlsx' → tab 'Database combine'.

Golden grain = 承批公司 × DICJ Code × 項目名稱 × Period(23/24/24_23SY/25/25_24SY/25_23SY) ×
Amount Type(報告/計劃/調整後) → Amount (單位 萬).

This script, per entity:
  1. dumps golden structure (承批公司→alias, #DICJ, #項目, DICJ 格式 sample),
  2. checks whether our kpi_report `dicj_code` col is populated + how many golden DICJ codes it matches,
  3. checks 項目名稱 match rate (golden name ⊂ our project col),
  4. project-level tie: per (project, bucket) golden vs ours — 報告 for 25-buckets, 調整後 for 24/23.
     (amount_mop ÷ 1e4 → 萬 to compare with golden.)

Run (Windows):  python scripts/inspect_dicj_match.py
Output: prints + results/inspect_dicj_match.txt  (paste back). Golden file is read-only, never committed.
"""
from __future__ import annotations
import sys, re
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
# 承批公司 (golden) → our alias, by substring (covers EN + 中文 變體)
NAME2ALIAS = [
    (["galaxy", "銀河"], "galaxy"),
    (["wynn", "永利"], "wynn"),
    (["mgm", "美高梅"], "mgm"),
    (["melco", "新濠", "摩珀斯", "影匯", "影滙", "studio city", "city of dreams", "cod"], "melco"),
    (["sjm", "澳娛", "葡京", "回力", "上葡京"], "sjm"),
    (["威尼斯", "金沙", "sands", "londoner", "倫敦人", "parisian", "巴黎人", "vml", "venetian"], "vml"),
]
DICJ_CANDS = ["dicj_code", "DICJ Code", "DICJ", "DICJ code", "dicj"]
PROJ_CANDS = ["Project Name", "Project", "SubProject_Name", "Project name - Amended", "Project_code",
              "項目名稱", "Name of Investment Project", "project", "Subproject"]


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


def _s(x):
    return str(x).strip() if x is not None else ""


def _num(s):
    return pd.to_numeric(pd.Series(s).astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0.0)


def main():
    L = ["# inspect_dicj_match — project-level tie vs golden 'Database combine'"]
    gpath = next((p for p in GCAND if p.exists()), None)
    if not gpath:
        L.append("!! golden 檔揾唔到 (HQ 投資方向_audit_0616.xlsx 喺 root/data/results)"); _w(L); return
    g = pd.read_excel(gpath, sheet_name=TAB, dtype=str)
    g.columns = [str(c).strip() for c in g.columns]
    L.append(f"golden={gpath.name}  tab={TAB!r}  rows={len(g):,}  cols={list(g.columns)}")
    c_ent = _col(g, ["承批公司"]); c_dicj = _col(g, ["DICJ Code", "DICJ"]); c_proj = _col(g, ["項目名稱"])
    c_per = _col(g, ["Period"]); c_typ = _col(g, ["Amount Type"]); c_amt = _col(g, ["Amount"])
    L.append(f"detected: 承批公司={c_ent!r} DICJ={c_dicj!r} 項目名稱={c_proj!r} Period={c_per!r} "
             f"Amount Type={c_typ!r} Amount={c_amt!r}")
    if not all([c_ent, c_dicj, c_proj, c_per, c_typ, c_amt]):
        L.append("!! 缺欄,改 detection"); _w(L); return

    g["_alias"] = g[c_ent].map(_alias)
    L.append(f"\n承批公司 → alias 對照:")
    for v, sub in g.groupby(c_ent):
        L.append(f"   {str(v)[:24]:24s} → {_alias(v)!r}  ({sub[c_proj].nunique()} 項目, {len(sub)} 行)")
    unmapped = sorted(set(g[g["_alias"].isna()][c_ent].astype(str)))
    if unmapped:
        L.append(f"   ⚠ 未對到 alias: {unmapped}")

    g["_amt"] = _num(g[c_amt])
    g["_per"] = g[c_per].map(_s)
    g["_typ"] = g[c_typ].map(_s)
    g["_dicj"] = g[c_dicj].map(_s)
    g["_proj"] = g[c_proj].map(_s)

    for alias, com in ENT.items():
        L.append(f"\n{'='*72}\n## {alias}")
        ga = g[g["_alias"] == alias]
        if not len(ga):
            L.append("   (golden 冇呢家)"); continue
        gdicj = set(d for d in ga["_dicj"].unique() if d)
        gname = set(n for n in ga["_proj"].unique() if n)
        L.append(f"   golden: {len(gdicj)} DICJ codes, {len(gname)} 項目名稱")
        L.append(f"   DICJ 格式 sample: {sorted(gdicj)[:8]}")

        p = ROOT / "data" / alias / "output" / f"{com}_kpi_report.parquet"
        if not p.exists():
            L.append(f"   (no kpi_report: {p.name})"); continue
        df = pd.read_parquet(p)
        dcol = _col(df, DICJ_CANDS)
        pcol = _col(df, PROJ_CANDS)
        odicj = set()
        if dcol:
            s = df[dcol].astype("string").fillna("").str.strip()
            odicj = set(x for x in s.unique() if x)
            L.append(f"   our dicj col={dcol!r}: {int(s.ne('').sum()):,}/{len(df):,} rows populated, {len(odicj)} distinct")
        else:
            L.append(f"   our dicj col = NONE (raw 冇 DICJ 欄)")
        oname = set()
        if pcol:
            s = df[pcol].astype("string").fillna("").str.strip()
            oname = set(x for x in s.unique() if x)
            L.append(f"   our project col={pcol!r}: {len(oname)} distinct")

        # match rates
        if gdicj:
            hitd = gdicj & odicj
            L.append(f"   ▶ DICJ match: {len(hitd)}/{len(gdicj)} golden DICJ 喺我哋 dicj_code ({len(hitd)/max(len(gdicj),1)*100:.0f}%)")
            miss = sorted(gdicj - odicj)[:8]
            if miss: L.append(f"      未 match DICJ: {miss}")
        if gname:
            hitn = gname & oname
            # also substring: golden name appears inside one of our project values
            sub_hit = 0
            for gn in (gname - oname):
                if any(gn and gn in on for on in oname):
                    sub_hit += 1
            L.append(f"   ▶ 名 match: exact {len(hitn)}/{len(gname)} ({len(hitn)/max(len(gname),1)*100:.0f}%) + substring 再加 {sub_hit}")
            missn = sorted(gname - oname)[:5]
            if missn: L.append(f"      未 exact-match 名 sample: {[n[:24] for n in missn]}")

        # project-level tie (報告 for 25-buckets, 調整後 for 24/23) — match by DICJ if我哋有,否則名
        amt_col = _col(df, ["amount_mop", "Val/COArea Crcy", "MOP Amt", "Amount - Amended",
                            "Debit minus Credit", "amount"])
        if amt_col is None or "report_period" not in df.columns:
            L.append("   (no amount/report_period — skip per-project tie)"); continue
        df["_w"] = pd.to_numeric(df[amt_col], errors="coerce").fillna(0.0) / 1e4   # MOP → 萬
        key_col = dcol if (dcol and len(gdicj & odicj) >= len(gname & oname)) else pcol
        gkey = "_dicj" if key_col == dcol else "_proj"
        L.append(f"   tie key = {key_col!r} (對 golden {gkey})")
        ours = (df.assign(_k=df[key_col].astype("string").fillna("").str.strip(),
                          _b=df["report_period"].astype("string").fillna(""))
                  .groupby(["_k", "_b"])["_w"].sum())
        def gtype(bk):
            return "報告" if str(bk).startswith("25") else "調整後"
        gg = ga[ga["_typ"].isin(["報告", "調整後"])].copy()
        gg["_use"] = gg["_per"].map(gtype)
        gsel = gg[gg["_typ"] == gg["_use"]]
        gold = gsel.groupby([gkey.replace("_", "_"), "_per"])["_amt"].sum()
        gold.index = gold.index.set_names(["_k", "_b"])
        joined = pd.concat([gold.rename("golden"), ours.rename("ours")], axis=1).fillna(0.0)
        joined["Δ"] = joined["ours"] - joined["golden"]
        tot_g = joined["golden"].sum(); tot_o = joined["ours"].sum()
        L.append(f"   ▶ project-bucket tie 總計: golden={tot_g:,.0f}萬  ours={tot_o:,.0f}萬  Δ={tot_o-tot_g:,.0f}萬")
        big = joined[joined["Δ"].abs() > 50].reindex(joined["Δ"].abs().sort_values(ascending=False).index).head(12)
        if len(big):
            L.append("   top |Δ| (key | bucket | golden | ours | Δ, 萬):")
            for (k, b), r in big.iterrows():
                L.append(f"      {str(k)[:30]:30s} {str(b):8s} {r['golden']:>10,.0f} {r['ours']:>10,.0f} {r['Δ']:>10,.0f}")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_dicj_match.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste back")


if __name__ == "__main__":
    main()
