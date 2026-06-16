r"""Per-(DICJ × bucket) side-by-side: golden (報告/計劃/調整後 三種全列) vs ours — 肉眼查 23 對唔上真因.

回答「兩邊都有 bucket 點解對唔上」：唔靠斷言, 攤開實數。每家出:
  1. bucket-level 總數對照 (golden 期望type vs ours) + 23+24_23SY 合併 check (睇係咪 SY 重疊切法唔同)
  2. 我哋 amount 究竟最似 golden 邊個 amount-type (報告/計劃/調整後) — 揭穿 amount-type 揀錯
  3. top per-(項目,bucket) gap → tie_detail_<ent>.tsv 畀你肉眼核

dicj 用我哋 kpi_report 嘅 dicj_code (sjm/vml/mgm 已 100% 填且驗證 100% 啱)。amt 只 check 數。

Run (Windows):  python scripts/inspect_tie_detail.py            # all 6
                python scripts/inspect_tie_detail.py vml sjm
"""
from __future__ import annotations
import sys
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


def _alias(s):
    sl = str(s).strip().lower()
    for kws, a in NAME2ALIAS:
        if any(k.lower() in sl for k in kws):
            return a
    return None


def _col(df, cands):
    low = {str(c).strip().lower(): c for c in df.columns}
    for c in cands:
        if c in df.columns: return c
        if str(c).strip().lower() in low: return low[str(c).strip().lower()]
    return None


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
    L = ["# inspect_tie_detail — golden(報告/計劃/調整後) vs ours, per (項目 × bucket), 肉眼查 23"]
    gpath = next((p for p in GCAND if p.exists()), None)
    if not gpath:
        L.append("!! golden 揾唔到"); _w(L); return
    g = pd.read_excel(gpath, sheet_name=TAB, dtype=str)
    g.columns = [str(c).strip() for c in g.columns]
    g["_a"] = g[next(c for c in g.columns if "承批" in c)].map(_alias)
    g["_d"] = g[next(c for c in g.columns if c.strip() in ("DICJ Code", "DICJ"))].astype(str).str.strip()
    g["_p"] = g[next(c for c in g.columns if c.strip() == "Period")].astype(str).str.strip()
    def _ntype(t):                                   # 簡繁/英文 amount-type 容錯 → canonical
        t = str(t).strip()
        cmap = {"報告": ["報告", "报告", "report", "actual", "實際", "实际"],
                "計劃": ["計劃", "计划", "plan", "budget", "預算", "预算"],
                "調整後": ["調整後", "调整后", "adjust", "adj"]}
        for canon, kws in cmap.items():
            if any(k.lower() in t.lower() for k in kws): return canon
        return t
    g["_t"] = g[next(c for c in g.columns if "Amount Type" in c)].astype(str).str.strip().map(_ntype)
    g["_amt"] = pd.to_numeric(g[next(c for c in g.columns if c.strip() == "Amount")].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0.0)
    outdir = ROOT / "results"; outdir.mkdir(exist_ok=True)

    for alias in want:
        com = ENT.get(alias)
        if not com: continue
        L.append(f"\n{'='*72}\n## {alias}")
        ga = g[g["_a"] == alias]
        L.append(f"   golden: {len(ga):,} 行  Σ_amt={ga['_amt'].sum():,.0f}  Period={sorted(ga['_p'].unique())}  Type={sorted(ga['_t'].unique())}")
        if len(ga) == 0:
            L.append("   !! golden 0 行 (alias 對唔到承批公司) — 睇 g['承批公司'].unique() 同 NAME2ALIAS"); continue
        # golden pivot: (dicj, period) × type → amount
        gp = ga.groupby(["_d", "_p", "_t"])["_amt"].sum().unstack("_t").fillna(0.0)
        for t in ("報告", "計劃", "調整後"):
            if t not in gp.columns: gp[t] = 0.0
        p = ROOT / "data" / alias / "output" / f"{com}_kpi_report.parquet"
        if not p.exists():
            L.append("   (no kpi_report)"); continue
        df = pd.read_parquet(p)
        dcol, amtc = _col(df, DICJ_CANDS), _conf_amt(df, com)
        if not dcol:
            L.append("   (no dicj_code col)"); continue
        dd = df[dcol].astype("string").fillna("").str.strip()
        rp = df["report_period"].astype(str).str.strip() if "report_period" in df.columns else pd.Series("", index=df.index)
        w = pd.to_numeric(df[amtc], errors="coerce").fillna(0.0) / 1e4
        ours = pd.DataFrame({"_d": dd.values, "_p": rp.values, "o": w.values}).groupby(["_d", "_p"])["o"].sum()

        # join
        idx = gp.index.union(ours.index)
        j = pd.DataFrame(index=idx)
        j["報告"] = gp["報告"].reindex(idx).fillna(0.0)
        j["計劃"] = gp["計劃"].reindex(idx).fillna(0.0)
        j["調整後"] = gp["調整後"].reindex(idx).fillna(0.0)
        j["ours"] = ours.reindex(idx).fillna(0.0)
        j = j.reset_index().rename(columns={"_d": "項目", "_p": "bucket"})
        j["expected"] = j.apply(lambda r: r["報告"] if str(r["bucket"]).startswith("25") else r["調整後"], axis=1)
        j["Δ_vs_expected"] = j["ours"] - j["expected"]
        # 我哋最似 golden 邊個 type
        def _best(r):
            d = {"報告": abs(r["ours"] - r["報告"]), "計劃": abs(r["ours"] - r["計劃"]), "調整後": abs(r["ours"] - r["調整後"])}
            return min(d, key=d.get)
        j["ours_最似"] = j.apply(_best, axis=1)

        # 1. bucket-level 總數對照
        bt = j.groupby("bucket")[["expected", "ours"]].sum()
        bt["Δ"] = bt["ours"] - bt["expected"]
        L.append("   ── bucket 總數 (golden expected vs ours) ──")
        for b, r in bt.iterrows():
            L.append(f"      {b:<8} golden={r['expected']:>14,.0f}  ours={r['ours']:>14,.0f}  Δ={r['Δ']:>12,.0f}")
        # SY 重疊 check: 23 vs 23+24_23SY ; 24 vs 24+25_24SY 等
        def _bsum(bs, col): return float(bt[col].reindex(bs).fillna(0.0).sum())
        for tgt, parts in (("23", ["23", "24_23SY", "25_23SY"]), ("24", ["24", "25_24SY"])):
            gtot = _bsum([tgt], "expected"); o_comb = _bsum(parts, "ours")
            if abs(gtot) > 1:
                L.append(f"      [SY check] golden {tgt} = {gtot:,.0f}  vs ours {'+'.join(parts)} = {o_comb:,.0f}  Δ={o_comb-gtot:,.0f}")

        # 2. amount-type 揀錯? 統計 ours 最似邊個 type (限 golden 有數嗰啲)
        nz = j[(j[["報告", "計劃", "調整後"]].abs().sum(axis=1) > 0)]
        L.append(f"   ── ours 最似 golden 邊個 type: {nz['ours_最似'].value_counts().to_dict()} (應該幾乎全部 '調整後' for 23/24)")

        # 3. top per-(項目,bucket) gap
        big = j[j["Δ_vs_expected"].abs() > 5].sort_values("Δ_vs_expected", key=lambda s: s.abs(), ascending=False)
        L.append(f"   ── top gap (|Δ|>5萬): {len(big)} 條 → tie_detail_{alias}.tsv  (頭 12):")
        for _, r in big.head(12).iterrows():
            L.append(f"      {str(r['項目'])[:14]:<14} {str(r['bucket']):<8} 報告={r['報告']:>11,.0f} 計劃={r['計劃']:>11,.0f} 調整後={r['調整後']:>11,.0f} | ours={r['ours']:>11,.0f} Δ={r['Δ_vs_expected']:>11,.0f}")
        big.round(1).to_csv(outdir / f"tie_detail_{alias}.tsv", sep="\t", index=False)
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_tie_detail.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)} + tie_detail_*.tsv  ← paste 返嚟")


if __name__ == "__main__":
    main()
