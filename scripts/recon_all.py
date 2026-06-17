r"""全 6 家 golden vs ours 對數總表：整體(bucket Σ) + 個別(by project) + 不能 fill(blank dicj)。

回答：①而家邊家 tie ②有冇行唔能 fill(dicj 空,落唔到 golden)③逐 project 同 golden 嘅整體+個別差異。

每家出：
  • DICJ 覆蓋：blank dicj 行數 + 佔額(萬) + %  ← 「不能 fill」就係呢啲
  • bucket Σ：golden(報告/調整後) vs ours，Δ  ← 整體差異
  • by-project：matched 但對唔上 / golden-only / ours-only（按金額排）  ← 個別差異

dicj 用 kpi_report 嘅 dicj_code（mgm 已經喺 pipeline remap 好）。amt 只 check 數。
sjm golden 項目37.1-6 → roll 項目37（project 級）。vml 24 golden 數字有問題(輸入錯)→ 會標。

Run (Windows):  python scripts/recon_all.py            # all 6
                python scripts/recon_all.py galaxy sjm
出 results/recon_all.txt + recon_all_<ent>_<bucket>.tsv  ← paste 返嚟
"""
from __future__ import annotations
import sys, re
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parent.parent
GCAND = [ROOT / "HQ 投資方向_audit_0616.xlsx", ROOT / "data" / "HQ 投資方向_audit_0616.xlsx"]
TAB = "Database combine"
ENT = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
       "vml": "company_4", "melco": "company_5", "mgm": "company_6"}
NAME2ALIAS = [(["galaxy", "銀河"], "galaxy"), (["wynn", "永利"], "wynn"), (["mgm", "美高梅"], "mgm"),
              (["melco", "新濠", "摩珀斯", "影匯", "影滙", "studio city", "city of dreams"], "melco"),
              (["sjm", "澳娛", "葡京", "回力", "上葡京"], "sjm"),
              (["威尼斯", "金沙", "sands", "londoner", "倫敦人", "parisian", "巴黎人", "venetian", "vml"], "vml")]
DICJ_CANDS = ["dicj_code", "DICJ Code", "DICJ", "dicj"]
PROJ_ROLLUP = {"sjm"}
NOTE = {("vml", "24"): "⚠ golden 24 數字輸入錯(我哋啱)"}
BUCKETS = ["25", "25_24SY", "25_23SY", "24", "24_23SY", "23"]


def _etype(b):
    return "報告" if str(b).startswith("25") else "調整後"


def _roll(alias, code):
    if alias in PROJ_ROLLUP:
        m = re.match(r"^(項目\d+)\.", str(code))
        if m: return m.group(1)
    return code


def _alias(s):
    sl = str(s).strip().lower()
    for kws, a in NAME2ALIAS:
        if any(k.lower() in sl for k in kws): return a
    return None


def _ntype(t):
    t = str(t)
    if any(k in t.lower() for k in ["報告", "报告", "report", "actual", "實際", "实际"]): return "報告"
    if any(k in t.lower() for k in ["計劃", "计划", "plan", "budget", "預算", "预算"]): return "計劃"
    if any(k in t.lower() for k in ["調整後", "调整后", "adjust", "adj"]): return "調整後"
    return t.strip()


def _col(df, cands):
    low = {str(c).strip().lower(): c for c in df.columns}
    for c in cands:
        if c in df.columns: return c
        if str(c).strip().lower() in low: return low[str(c).strip().lower()]
    return None


def _conf_amt(df, com):
    try:
        a = (yaml.safe_load((ROOT / "conf" / com / "parameters.yml").read_text(encoding="utf-8")).get("columns") or {}).get("amount", "")
        if a and a in df.columns: return a
    except Exception: pass
    return _col(df, ["Val/COArea Crcy", "MOP Amt", "Amount - Amended", "Debit minus Credit",
                     "Entry Voucher Amount/ Expense Amount", "amount_mop", "amount"])


def main():
    want = [a.lower() for a in sys.argv[1:]] or list(ENT)
    L = ["# recon_all — 6 家 golden vs ours：覆蓋(不能fill) + bucket Σ(整體) + by-project(個別)"]
    gpath = next((p for p in GCAND if p.exists()), None)
    if not gpath:
        L.append("!! golden 揾唔到"); _w(L); return
    g = pd.read_excel(gpath, sheet_name=TAB, dtype=str)
    g.columns = [str(c).strip() for c in g.columns]
    g["_a"] = g[next(c for c in g.columns if "承批" in c)].map(_alias)
    g["_d"] = g[next(c for c in g.columns if c.strip() in ("DICJ Code", "DICJ"))].astype(str).str.strip()
    g["_p"] = g[next(c for c in g.columns if c.strip() == "Period")].astype(str).str.strip()
    g["_t"] = g[next(c for c in g.columns if "Amount Type" in c)].map(_ntype)
    g["_amt"] = pd.to_numeric(g[next(c for c in g.columns if c.strip() == "Amount")].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0.0)
    outdir = ROOT / "results"; outdir.mkdir(exist_ok=True)

    overall = []
    for alias in want:
        com = ENT.get(alias)
        if not com: continue
        L.append(f"\n{'='*72}\n## {alias}")
        ga = g[g["_a"] == alias].copy()
        if alias in PROJ_ROLLUP:
            ga["_d"] = ga["_d"].map(lambda c: _roll(alias, c))
        p = ROOT / "data" / alias / "output" / f"{com}_kpi_report.parquet"
        if not p.exists():
            L.append("   (no kpi_report parquet — 未跑)"); continue
        df = pd.read_parquet(p)
        dcol, amtc = _col(df, DICJ_CANDS), _conf_amt(df, com)
        dd = df[dcol].astype("string").fillna("").str.strip() if dcol else pd.Series("", index=df.index)
        if alias in PROJ_ROLLUP:
            dd = dd.map(lambda c: _roll(alias, c))
        w = pd.to_numeric(df[amtc], errors="coerce").fillna(0.0) / 1e4 if amtc else pd.Series(0.0, index=df.index)
        rp = df["report_period"].astype(str).str.strip() if "report_period" in df.columns else pd.Series("", index=df.index)

        # ── ① DICJ 覆蓋（不能 fill = blank dicj）──
        blank = dd.eq("")
        nblank = int(blank.sum()); ntot = len(df)
        blank_amt = float(w[blank].sum())
        L.append(f"   ① DICJ 覆蓋: filled {ntot - nblank:,}/{ntot:,} ({(ntot-nblank)/max(ntot,1)*100:.1f}%)  | blank {nblank:,} 行 = {blank_amt:,.0f}萬 唔能 fill golden")
        if nblank:
            bb = pd.DataFrame({"_p": rp[blank].values, "o": w[blank].values}).groupby("_p")["o"].sum()
            L.append("      blank by bucket(萬): " + ", ".join(f"{k}={v:,.0f}" for k, v in bb.items() if abs(v) > 0.5))

        # ── ② bucket Σ（整體差異）──
        L.append("   ② bucket 總數 (golden vs ours, 萬):")
        ent_ties = True
        for b in BUCKETS:
            et = _etype(b)
            gsum = float(ga[(ga["_p"] == b) & (ga["_t"] == et)]["_amt"].sum())
            osum = float(w[rp == b].sum())
            if abs(gsum) < 1 and abs(osum) < 1: continue
            d = osum - gsum
            note = "  " + NOTE.get((alias, b), "") if (alias, b) in NOTE else ""
            flag = "" if abs(d) <= 5 else ("  ←Δ" if not note else "")
            if abs(d) > 5 and not note: ent_ties = False
            L.append(f"      {b:<8} golden={gsum:>12,.0f}  ours={osum:>12,.0f}  Δ={d:>10,.0f}{note}{flag}")

        # ── ③ by-project（個別差異）──
        gp = ga.groupby(["_d", "_p", "_t"])["_amt"].sum().unstack("_t").fillna(0.0)
        for t in ("報告", "計劃", "調整後"):
            if t not in gp.columns: gp[t] = 0.0
        oj = pd.DataFrame({"_d": dd[~blank].values, "_p": rp[~blank].values, "o": w[~blank].values}).groupby(["_d", "_p"])["o"].sum()
        big_n = gonly_n = oonly_n = 0
        rows_all = []
        for b in BUCKETS:
            et = _etype(b)
            gb = gp[et][gp.index.get_level_values("_p") == b] if isinstance(gp.index, pd.MultiIndex) else pd.Series(dtype=float)
            gb = gb.groupby(level="_d").sum() if len(gb) else pd.Series(dtype=float)
            gb = gb[gb.abs() > 1e-9]
            ob = oj[oj.index.get_level_values("_p") == b].groupby(level="_d").sum() if len(oj) else pd.Series(dtype=float)
            codes = sorted(set(gb.index) | set(ob.index))
            for c in codes:
                gv, ov = float(gb.get(c, 0.0)), float(ob.get(c, 0.0))
                st = "match" if (c in gb.index and c in ob.index) else ("golden_only" if c in gb.index else "ours_only")
                rows_all.append({"bucket": b, "code": c, "golden": round(gv, 1), "ours": round(ov, 1), "Δ": round(ov - gv, 1), "status": st})
        rdf = pd.DataFrame(rows_all)
        if len(rdf):
            mism = rdf[(rdf.status == "match") & (rdf["Δ"].abs() > 5)].sort_values("Δ", key=lambda s: s.abs(), ascending=False)
            gonly = rdf[(rdf.status == "golden_only") & (rdf["golden"].abs() > 5)].sort_values("golden", key=lambda s: s.abs(), ascending=False)
            oonly = rdf[(rdf.status == "ours_only") & (rdf["ours"].abs() > 5)].sort_values("ours", key=lambda s: s.abs(), ascending=False)
            big_n, gonly_n, oonly_n = len(mism), len(gonly), len(oonly)
            L.append(f"   ③ by-project (|Δ|>5萬): mismatch={big_n}  golden_only={gonly_n}(Σ{gonly['golden'].sum():,.0f})  ours_only={oonly_n}(Σ{oonly['ours'].sum():,.0f})")
            for tag, sub in (("對唔上", mism), ("golden有我冇", gonly), ("我有golden冇", oonly)):
                if len(sub):
                    L.append(f"      ── {tag} (頭6) ──")
                    for _, r in sub.head(6).iterrows():
                        L.append(f"         {r['bucket']:<8} {str(r['code'])[:12]:<12} golden={r['golden']:>10,.0f} ours={r['ours']:>10,.0f} Δ={r['Δ']:>9,.0f}")
            rdf.round(1).to_csv(outdir / f"recon_all_{alias}.tsv", sep="\t", index=False)
        # verdict by AMOUNT not count: Σ|Δ| of all by-project diffs ÷ entity total (exclude known golden-error buckets)
        err_bk = {b for (a, b) in NOTE if a == alias}
        diff_amt = float(rdf[~rdf["bucket"].isin(err_bk)]["Δ"].abs().sum()) if len(rdf) else 0.0
        gtot_ent = sum(float(ga[(ga["_p"] == b) & (ga["_t"] == _etype(b))]["_amt"].sum()) for b in BUCKETS)
        pct = diff_amt / max(abs(gtot_ent), 1) * 100
        verdict = "✅ TIE" if pct < 0.5 else ("◐ 近tie" if pct < 1.5 else "✗ 有差")
        if err_bk: verdict += " *(golden側錯)"
        overall.append((alias, verdict, nblank, blank_amt, diff_amt, pct))

    L.append(f"\n{'='*72}\n## 總結 (個別差異按金額%判，已扣 golden-側錯 bucket)")
    for a, v, nb, ba, da, pc in overall:
        L.append(f"   {a:<8} {v:<18} blank={nb:,}行/{ba:,.0f}萬  個別Σ|Δ|={da:,.0f}萬 ({pc:.2f}%)")
    _w(L)


def _w(L):
    out = ROOT / "results" / "recon_all.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)} + recon_all_<ent>.tsv  ← paste 返嚟")


if __name__ == "__main__":
    main()
