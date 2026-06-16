r"""MGM 專用 DICJ reconciliation：golden 用邊個 code、我哋出邊個 code，逐 bucket 攤開對齊。

背景（user 對照表 + inspect_ptteam）：
  • golden mgm 用 *consolidated* 項目N（Master_code，去零）；23/24 gaming = 博彩項目N。
  • 我哋 raw / report 用 *split* Project_code：項目0NN-OPEX / -CAPEX / -製作期、項目CAPEX-N，
    24 gaming 又 user 自己加咗「博彩」前綴 → 同 golden 對唔到（25 拆碼 + 24 博彩前綴）。

呢個 script 唔砌嘢、唔改 data，淨係攤開實數畀你 + 我肉眼核：
  1. 確認 amount col 啱（bucket 總數要 = inspect_tie_detail：25≈186,143 / 24≈149,761 / 23≈95,009 萬）
  2. 試一條 normalize（去零 + 去 -OPEX/-CAPEX/-製作期 尾），睇 25 拆碼 roll 返 parent 之後對唔對得返 golden
  3. 逐 bucket 出三張表：matched(|Δ|) / golden-only / ours-only（按金額排，方便用金額配對揾 gaming 25 點 map）
  4. 另讀 mgm_2025.xlsx[Project_all] 的 Project_code→Master_code，試解 ours-only（尤其 項目CAPEX-N）

Run (Windows):  python scripts/recon_mgm.py
出 results/recon_mgm.txt + recon_mgm_<bucket>.tsv  ← paste 返嚟
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
COM, ALIAS = "company_6", "mgm"
NAME_KWS = ["mgm", "美高梅"]
BUCKETS = ["25", "25_24SY", "25_23SY", "24", "24_23SY", "23"]


def _ndzero(c):
    """項目019 → 項目19 ; 博彩項目05 → 博彩項目5 (去 code 前綴後嘅 leading zero)."""
    c = str(c).strip()
    m = re.match(r"^(博彩項目|項目)0*(\d.*)$", c)
    return m.group(1) + m.group(2) if m else c


def norm_ours(code, strip_bochai=False):
    """我哋 split code → golden consolidated 試法：去零 + 去 -OPEX/-CAPEX/-製作期 尾."""
    c = _ndzero(code)
    c = re.sub(r"^(項目\d+)[-－](?:OPEX|CAPEX).*$", r"\1", c)   # 項目19-OPEX / 19-CAPEX → 項目19
    c = re.sub(r"^(項目\d+)製作期.*$", r"\1", c)                # 項目17製作期 → 項目17
    if strip_bochai:
        c = re.sub(r"^博彩(項目\d+)$", r"\1", c)                # 24 gaming: 博彩項目14 → 項目14
    return c


def _etype(bucket):
    return "報告" if str(bucket).startswith("25") else "調整後"


def _ntype(t):
    t = str(t)
    if any(k in t for k in ["報告", "报告", "report"]): return "報告"
    if any(k in t for k in ["調整後", "调整后", "adjust", "adj"]): return "調整後"
    if any(k in t for k in ["計劃", "计划", "plan", "budget"]): return "計劃"
    return t.strip()


def _conf_amt(df, com):
    try:
        a = (yaml.safe_load((ROOT / "conf" / com / "parameters.yml").read_text(encoding="utf-8")).get("columns") or {}).get("amount", "")
        if a and a in df.columns: return a
    except Exception: pass
    for c in ["Debit minus Credit", "amount_mop", "amount", "MOP Amt", "Amount - Amended", "Val/COArea Crcy"]:
        if c in df.columns: return c
    return None


def _master_map():
    """讀 mgm_2025.xlsx[Project_all] → {Project_code(去零): Master_code(去零)}，解 ours-only."""
    f = ROOT / "data" / ALIAS / "raw" / "mgm_2025.xlsx"
    if not f.exists(): return {}, "（冇 mgm_2025.xlsx）"
    try:
        d = pd.read_excel(f, sheet_name="Project_all", dtype=str)
    except Exception as e:
        return {}, f"（讀唔到 Project_all: {e}）"
    cols = list(d.columns)
    pc = next((c for c in cols if str(c).strip() == "Project_code"), None)
    mc = next((c for c in cols if str(c).strip() == "Master_code"), None)
    if not pc or not mc: return {}, f"（Project_all 揾唔到 Project_code/Master_code，cols={cols[:8]}）"
    m = {}
    for _, r in d[[pc, mc]].dropna(how="all").iterrows():
        k, v = str(r[pc]).strip(), str(r[mc]).strip()
        if k and k.lower() != "nan" and v and v.lower() not in ("nan", "#n/a"):
            m[_ndzero(k)] = _ndzero(v)
    return m, f"（{len(m)} 條 Project_code→Master_code）"


def main():
    L = ["# recon_mgm — golden code vs ours，逐 bucket 攤開（去零＋去 -OPEX/-CAPEX/-製作期 尾）"]
    gpath = next((p for p in GCAND if p.exists()), None)
    if not gpath:
        L.append("!! golden 揾唔到"); _w(L); return
    g = pd.read_excel(gpath, sheet_name=TAB, dtype=str)
    g.columns = [str(c).strip() for c in g.columns]
    g = g[g[next(c for c in g.columns if "承批" in c)].astype(str).str.lower().map(
        lambda s: any(k.lower() in s for k in NAME_KWS))].copy()
    g["_d"] = g[next(c for c in g.columns if c.strip() in ("DICJ Code", "DICJ"))].astype(str).str.strip().map(_ndzero)
    g["_p"] = g[next(c for c in g.columns if c.strip() == "Period")].astype(str).str.strip()
    g["_t"] = g[next(c for c in g.columns if "Amount Type" in c)].map(_ntype)
    g["_amt"] = pd.to_numeric(g[next(c for c in g.columns if c.strip() == "Amount")].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0.0)

    p = ROOT / "data" / ALIAS / "output" / f"{COM}_kpi_report.parquet"
    if not p.exists():
        L.append("   (no kpi_report parquet)"); _w(L); return
    df = pd.read_parquet(p)
    amtc = _conf_amt(df, COM)
    dd_raw = df["dicj_code"].astype("string").fillna("").str.strip()
    rp = df["report_period"].astype(str).str.strip()
    w = pd.to_numeric(df[amtc], errors="coerce").fillna(0.0) / 1e4
    L.append(f"   amount col = '{amtc}'   (bucket 總數要對返 inspect_tie_detail: 25≈186,143 / 24≈149,761 / 23≈95,009)")

    mmap, mnote = _master_map()
    L.append(f"   Project_all master-map {mnote}")
    outdir = ROOT / "results"; outdir.mkdir(exist_ok=True)

    for b in BUCKETS:
        et = _etype(b)
        gb = g[(g["_p"] == b) & (g["_t"] == et)].groupby("_d")["_amt"].sum()
        gb = gb[gb.abs() > 1e-9]
        mask = (rp == b)
        ob_raw = pd.Series(w[mask].values, index=dd_raw[mask].values).groupby(level=0).sum()
        # normalize ours：去零＋去尾（連 24 博彩前綴一齊試）
        sb = "博彩" if b in ("24", "24_23SY") else ""   # 24 gaming user 加咗博彩前綴 → 試 strip
        nidx = [norm_ours(c, strip_bochai=bool(sb)) for c in ob_raw.index]
        ob = ob_raw.copy(); ob.index = nidx; ob = ob.groupby(level=0).sum()
        if gb.empty and ob.empty:
            continue
        gtot, otot = float(gb.sum()), float(ob.sum())
        L.append(f"\n{'='*72}\n## bucket {b}  (golden={et})   golden Σ={gtot:,.0f}  ours(normed) Σ={otot:,.0f}  Δ={otot-gtot:,.0f}")
        codes = sorted(set(gb.index) | set(ob.index))
        rows = []
        for c in codes:
            gv, ov = float(gb.get(c, 0.0)), float(ob.get(c, 0.0))
            rows.append({"code": c, "golden": round(gv, 1), "ours": round(ov, 1), "Δ": round(ov - gv, 1),
                         "status": "match" if (c in gb.index and c in ob.index) else ("golden_only" if c in gb.index else "ours_only")})
        rdf = pd.DataFrame(rows)
        matched = rdf[rdf.status == "match"]
        gonly = rdf[rdf.status == "golden_only"].sort_values("golden", key=lambda s: s.abs(), ascending=False)
        oonly = rdf[rdf.status == "ours_only"].sort_values("ours", key=lambda s: s.abs(), ascending=False)
        bigmatch = matched[matched["Δ"].abs() > 5].sort_values("Δ", key=lambda s: s.abs(), ascending=False)
        L.append(f"   matched={len(matched)} (|Δ|>5萬:{len(bigmatch)})   golden_only={len(gonly)} Σ={gonly['golden'].sum():,.0f}   ours_only={len(oonly)} Σ={oonly['ours'].sum():,.0f}")
        if len(bigmatch):
            L.append("   ── matched 但對唔上 (頭8) ──")
            for _, r in bigmatch.head(8).iterrows():
                L.append(f"      {r['code']:<14} golden={r['golden']:>11,.0f}  ours={r['ours']:>11,.0f}  Δ={r['Δ']:>10,.0f}")
        if len(gonly):
            L.append("   ── golden 有、我哋冇 (頭12，金額排) ──")
            for _, r in gonly.head(12).iterrows():
                L.append(f"      {r['code']:<14} golden={r['golden']:>11,.0f}")
        if len(oonly):
            L.append("   ── 我哋有、golden 冇 = 未 map (頭12，金額排) ── [master-map 建議]")
            for _, r in oonly.head(12).iterrows():
                sug = mmap.get(r["code"], "—")
                L.append(f"      {r['code']:<14} ours={r['ours']:>11,.0f}   → master={sug}")
        rdf.to_csv(outdir / f"recon_mgm_{b}.tsv", sep="\t", index=False)
    _w(L)


def _w(L):
    out = ROOT / "results" / "recon_mgm.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)} + recon_mgm_<bucket>.tsv  ← paste 返嚟")


if __name__ == "__main__":
    main()
