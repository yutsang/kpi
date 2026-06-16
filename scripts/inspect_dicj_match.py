r"""Project-level tie vs the NEW golden 'HQ 投資方向_audit_0616.xlsx' → tab 'Database combine'.

Golden grain = 承批公司 × DICJ Code × 項目名稱 × Period(23/24/24_23SY/25/25_24SY/25_23SY) ×
Amount Type(報告/計劃/調整後) → Amount (單位 萬).

Per entity:
  1. golden structure (承批公司→alias, #DICJ, #項目, DICJ 格式 sample),
  2. our dicj_code populated %,
  3. match rate for 3 key-pairings — our_dicj↔golden_DICJ / our_proj↔golden_DICJ / our_proj↔golden_name
     — pick the BEST as the tie key (covers galaxy=DICJ, mgm=Project_code↔DICJ, sjm/melco=name),
  4. project-level tie by best key: 報告 for 25-buckets, 調整後 for 24/23. amount = conf columns.amount
     (NOT amount_mop — that col is partial for the unified-23 entities). amount ÷ 1e4 → 萬.

Run (Windows):  python scripts/inspect_dicj_match.py
Output: prints + results/inspect_dicj_match.txt  (paste back). Golden file is read-only, never committed.
"""
from __future__ import annotations
import sys
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


def _num(s):
    return pd.to_numeric(pd.Series(s).astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0.0)


def _conf_amt(df, com):
    try:
        cfg = yaml.safe_load((ROOT / "conf" / com / "parameters.yml").read_text(encoding="utf-8"))
        a = (cfg.get("columns") or {}).get("amount", "")
        c = _col(df, [a]) if a else None
        if c:
            return c
    except Exception:
        pass
    return _col(df, ["Val/COArea Crcy", "MOP Amt", "Amount - Amended", "Debit minus Credit",
                     "Reported Amount(MOP)", "Entry Voucher Amount/ Expense Amount", "amount_mop", "amount"])


def main():
    L = ["# inspect_dicj_match v2 — project-level tie vs golden 'Database combine'"]
    gpath = next((p for p in GCAND if p.exists()), None)
    if not gpath:
        L.append("!! golden 揾唔到"); _w(L); return
    g = pd.read_excel(gpath, sheet_name=TAB, dtype=str)
    g.columns = [str(c).strip() for c in g.columns]
    c_ent, c_dicj, c_proj = _col(g, ["承批公司"]), _col(g, ["DICJ Code", "DICJ"]), _col(g, ["項目名稱"])
    c_per, c_typ, c_amt = _col(g, ["Period"]), _col(g, ["Amount Type"]), _col(g, ["Amount"])
    g["_alias"] = g[c_ent].map(_alias)
    g["_amt"] = _num(g[c_amt]); g["_per"] = g[c_per].astype(str).str.strip()
    g["_typ"] = g[c_typ].astype(str).str.strip()
    g["_dicj"] = g[c_dicj].astype(str).str.strip(); g["_proj"] = g[c_proj].astype(str).str.strip()
    L.append(f"golden rows={len(g):,}  承批公司→alias: " +
             "  ".join(f"{v}→{_alias(v)}" for v in g[c_ent].dropna().unique()))

    for alias, com in ENT.items():
        L.append(f"\n{'='*72}\n## {alias}")
        ga = g[g["_alias"] == alias]
        if not len(ga):
            L.append("   (golden 冇)"); continue
        gdicj = set(d for d in ga["_dicj"].unique() if d and d != "nan")
        gname = set(n for n in ga["_proj"].unique() if n and n != "nan")
        L.append(f"   golden: {len(gdicj)} DICJ ({sorted(gdicj)[:5]}…), {len(gname)} 項目名稱")
        p = ROOT / "data" / alias / "output" / f"{com}_kpi_report.parquet"
        if not p.exists():
            L.append("   (no kpi_report)"); continue
        df = pd.read_parquet(p)
        dcol, pcol = _col(df, DICJ_CANDS), _col(df, PROJ_CANDS)
        od = df[dcol].astype("string").fillna("").str.strip() if dcol else pd.Series("", index=df.index)
        op = df[pcol].astype("string").fillna("").str.strip() if pcol else pd.Series("", index=df.index)
        L.append(f"   our dicj={dcol!r} ({int(od.ne('').sum()):,}/{len(df):,} filled) | proj={pcol!r}")

        # try 3 key-pairings; pick the one matching most golden keys
        cand = []
        if dcol: cand.append(("our_dicj↔gDICJ", od, set(od)-{''}, gdicj))
        if pcol:
            cand.append(("our_proj↔gDICJ", op, set(op)-{''}, gdicj))
            cand.append(("our_proj↔g名", op, set(op)-{''}, gname))
        best = None
        for nm, oser, oset, gset in cand:
            hit = len(oset & gset)
            L.append(f"   {nm}: {hit}/{len(gset)} golden keys 命中 ({hit/max(len(gset),1)*100:.0f}%)")
            if best is None or hit > best[0]:
                best = (hit, nm, oser, ("_dicj" if gset is gdicj else "_proj"))
        if best is None:
            L.append("   (冇 key 可配)"); continue
        _, bnm, oser, gkey = best
        L.append(f"   ▶ 用 key 配對: {bnm}")

        amt = _conf_amt(df, com)
        if amt is None or "report_period" not in df.columns:
            L.append("   (no amount/report_period)"); continue
        w = pd.to_numeric(df[amt], errors="coerce").fillna(0.0) / 1e4
        L.append(f"   amount 欄={amt!r}  ours 總額={w.sum():,.0f}萬")
        ours = pd.DataFrame({"_k": oser.values, "_b": df["report_period"].astype(str).values, "_w": w.values}) \
                 .groupby(["_k", "_b"])["_w"].sum()
        gg = ga.copy()
        gg["_use"] = gg["_per"].map(lambda b: "報告" if str(b).startswith("25") else "調整後")
        gsel = gg[gg["_typ"] == gg["_use"]]
        gold = gsel.groupby([gkey, "_per"])["_amt"].sum()
        gold.index = gold.index.set_names(["_k", "_b"])
        j = pd.concat([gold.rename("g"), ours.rename("o")], axis=1).fillna(0.0)
        j["d"] = j["o"] - j["g"]
        # matched-key tie (exclude keys absent from one side → those are name-format / unmatched noise)
        matched = j[(j["g"] != 0) & (j["o"] != 0)]
        L.append(f"   ▶ 總: golden={j['g'].sum():,.0f}  ours={j['o'].sum():,.0f}  Δ={j['d'].sum():,.0f}萬  "
                 f"| 已配對-key Δ={matched['d'].sum():,.0f}萬 ({len(matched)} key-bucket)")
        big = matched.reindex(matched["d"].abs().sort_values(ascending=False).index).head(8)
        if len(big):
            L.append("   top 已配對 |Δ| (key|bucket|g|o|Δ 萬):")
            for (k, b), r in big.iterrows():
                L.append(f"      {str(k)[:26]:26s} {str(b):8s} {r['g']:>9,.0f} {r['o']:>9,.0f} {r['d']:>8,.0f}")
        only_g = j[(j["g"] != 0) & (j["o"] == 0)]
        L.append(f"   ⚠ golden 有但我哋 key 對唔到: {len(only_g)} key-bucket, Σgolden={only_g['g'].sum():,.0f}萬 "
                 f"(= 名/code 對唔上嗰啲，要補 DICJ / 對名)")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_dicj_match.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste back")


if __name__ == "__main__":
    main()
