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
try:
    from tqdm import tqdm
except Exception:
    def tqdm(x, **k): return x

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
        from collections import defaultdict as _dd
        ga = g[g["_a"] == alias]
        # golden 報告/調整後 視角 (25→報告, 其餘→調整後): 反填同對數都用呢個
        gg = ga[ga["_t"].isin(["報告", "調整後"])].copy()
        gg["_use"] = gg["_p"].map(lambda b: "報告" if str(b).startswith("25") else "調整後")
        gsel = gg[gg["_t"] == gg["_use"]].copy()
        gsel["_nn"] = gsel["_n"].map(_norm); gsel["_pp"] = gsel["_p"].astype(str).str.strip()

        # ── PERIOD-AWARE golden map: (中文名, period) → DICJ (用行自己年份解 generic「其他」桶歧義) ──
        combo, rawbyname = _dd(set), {}
        for _, r in gsel.iterrows():
            _n, _d = r["_nn"], str(r["_d"]).strip()
            if _n and _d and _d != "nan":
                combo[(_n, r["_pp"])].add(_d); rawbyname.setdefault(_n, str(r["_n"]).strip())
        pa = {k: next(iter(v)) for k, v in combo.items() if len(v) == 1}     # (名,period) 唯一 → 填
        amb_pa = {k: v for k, v in combo.items() if len(v) > 1}              # 連 period 都歧義 → 唔填
        names_by_p = _dd(list)
        for (n, pp) in pa: names_by_p[pp].append(n)
        for pp in names_by_p: names_by_p[pp].sort(key=len, reverse=True)
        _gamt = gsel.groupby(["_nn", "_pp"])["_amt"].sum()
        amb_pa_amt = float(sum(abs(_gamt.get(k, 0.0)) for k in amb_pa))

        p = ROOT / "data" / alias / "output" / f"{com}_kpi_report.parquet"
        if not p.exists():
            L.append("   (no kpi_report)"); continue
        df = pd.read_parquet(p)
        dcol, pcol, amtc = _col(df, DICJ_CANDS), _col(df, PROJ_CANDS), _conf_amt(df, com)
        ex = df[dcol].astype("string").fillna("").str.strip() if dcol else pd.Series("", index=df.index)
        nm = df[pcol].astype("string").fillna("").str.strip() if pcol else pd.Series("", index=df.index)
        w = pd.to_numeric(df[amtc], errors="coerce").fillna(0.0) / 1e4 if amtc else pd.Series(0.0, index=df.index)
        rp = (df["report_period"].astype(str) if "report_period" in df.columns else pd.Series("", index=df.index)).str.strip()

        # period-aware 子字串反填: 行 period 嗰組 golden 中文名, 最長 substring → DICJ (cache per (name,period))
        _cache = {}
        def _subpa(name, period):
            k = (name, period)
            if k in _cache: return _cache[k]
            nn = _norm(name); hit = ""
            for gn in names_by_p.get(period, ()):
                if gn in nn: hit = pa[(gn, period)]; break
            _cache[k] = hit; return hit
        _keys = list(zip(nm.values, rp.values))
        for k in tqdm(set(_keys), desc=f"{alias} period-aware fill", leave=False):
            _subpa(k[0], k[1])
        sub_pred = pd.Series([_cache[k] for k in _keys], index=df.index)
        fd = ex.where(ex.ne(""), sub_pred)
        hw = pd.Series("unmatched", index=df.index).mask(sub_pred.ne(""), "name_substr").mask(ex.ne(""), "existing")

        tot = w.abs().sum(); cov_ex = w.abs()[ex.ne("")].sum(); cov_fd = w.abs()[fd.ne("")].sum()
        L.append(f"   amount={amtc!r}  Σ|amt|={tot:,.0f}萬")
        L.append(f"   PERIOD-AWARE map: {len(pa):,} (名,期)→唯一DICJ | 連period都仲歧義 {len(amb_pa)} 涉 {amb_pa_amt:,.0f}萬")
        L.append(f"   coverage by amount: existing {cov_ex/max(tot,1)*100:.0f}%  →  period-aware filled {cov_fd/max(tot,1)*100:.0f}%")
        L.append(f"   resolve how: {hw.value_counts().to_dict()}")

        # ── 驗證: existing 真值 vs period-aware 預測一致率 (填得啱唔啱) ──
        _same = pd.Series(sub_pred.values == ex.values, index=df.index)
        _gt = ex.ne("") & sub_pred.ne(""); _ag = _gt & _same
        n_gt, n_pred, a_cnt = int(ex.ne("").sum()), int(_gt.sum()), int(_ag.sum())
        wp, wa = w.abs()[_gt].sum(), w.abs()[_ag].sum()
        if n_pred:
            L.append(f"   ✅ 驗證(existing真值): 命中 {n_pred:,}/{n_gt:,}; 一致 {a_cnt:,}={a_cnt/n_pred*100:.1f}% (金額 {wa/max(wp,1)*100:.1f}%)  mismatch→dicj_verify_{alias}.tsv")
            _mm = _gt & ~_same
            if int(_mm.sum()):
                pd.DataFrame({"name": nm[_mm].values, "period": rp[_mm].values, "existing": ex[_mm].values, "pred": sub_pred[_mm].values}).drop_duplicates().head(200).to_csv(outdir / f"dicj_verify_{alias}.tsv", sep="\t", index=False)
        else:
            L.append(f"   驗證: existing 行零命中 (全自映/code-prefix, 正常)")

        # tie vs golden, before(existing) vs after(period-aware)
        gold = gsel.groupby(["_d", "_pp"])["_amt"].sum(); gold.index = gold.index.set_names(["_k", "_b"])
        for tag, key in (("BEFORE(existing)", ex), ("AFTER(period-aware)", fd)):
            o = pd.DataFrame({"_k": key.values, "_b": rp.values, "_w": w.values}).groupby(["_k", "_b"])["_w"].sum()
            j = pd.concat([gold.rename("g"), o.rename("o")], axis=1).fillna(0.0)
            m = j[(j["g"] != 0) & (j["o"] != 0)]; unmatched = j[(j["g"] != 0) & (j["o"] == 0)]["g"].sum()
            L.append(f"   {tag}: 已配對 Δ={m['o'].sum()-m['g'].sum():,.0f}萬 ({len(m)} key)  golden 對唔到={unmatched:,.0f}萬")

        # residual breakdown (AFTER)
        o_fd = pd.DataFrame({"_k": fd.values, "_b": rp.values, "_w": w.values}).groupby(["_k", "_b"])["_w"].sum()
        jj = pd.concat([gold.rename("g"), o_fd.rename("o")], axis=1).fillna(0.0)
        our_d = set(x for x in fd.unique() if str(x).strip()); gold_d = set(gold.index.get_level_values("_k"))
        miss_d = sorted(gold_d - our_d)
        gap = jj[(jj["g"] != 0) & (jj["o"] == 0)].reset_index().sort_values("g", key=lambda s: s.abs(), ascending=False)
        gap_miss = gap[gap["_k"].isin(miss_d)]["g"].abs().sum(); gap_buck = gap[~gap["_k"].isin(miss_d)]["g"].abs().sum()
        L.append(f"   殘差拆解: 缺dicj={gap_miss:,.0f}萬 ({len(miss_d)} code, e.g.{miss_d[:8]}) | bucket對唔上={gap_buck:,.0f}萬")
        gap.head(200).to_csv(outdir / f"dicj_goldgap_{alias}.tsv", sep="\t", index=False)

        # period-aware namemap (中文名 \t period \t DICJ) 供 step0 反填用
        pd.DataFrame([(rawbyname[n], pp, d) for (n, pp), d in pa.items()]).to_csv(
            outdir / f"dicj_namemap_{alias}.tsv", sep="\t", index=False, header=False)

        # lookup + unmatched
        lk = pd.DataFrame({"subproject": nm, "period": rp, "dicj": fd, "how": hw}).drop_duplicates(["subproject", "period"])
        lk.sort_values(["how", "subproject"]).to_csv(outdir / f"dicj_lookup_{alias}.tsv", sep="\t", index=False)
        un = lk[lk["dicj"] == ""].copy()
        if len(un):
            un["amt萬"] = un["subproject"].map(lambda x: round(w.abs()[nm.eq(x)].sum(), 1))
            un.sort_values("amt萬", ascending=False)[["subproject", "period", "amt萬"]].head(400).to_csv(
                outdir / f"dicj_unmatched_{alias}.tsv", sep="\t", index=False)
    _w(L)


def _w(L):
    out = ROOT / "results" / "build_dicj_fill.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)} + dicj_lookup_*/dicj_unmatched_*  ← paste back")


if __name__ == "__main__":
    main()
