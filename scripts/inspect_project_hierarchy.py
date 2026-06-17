r"""inspect_project_hierarchy — 為項目組要求嘅 hierarchy 做準備：
   project名 = golden DICJ 名稱（by dicj_code）｜ subproject = 最細 level 名 + subproject_code = 細 code。

每家出：
  ① golden dicj→名 覆蓋：我哋 dicj_code 有幾多 map 到 golden 名（mapped / unmapped 數 + Σ）
  ② unmapped / blank dicj 行：dicj + project + subproject 名 + Σ（睇點按名 map 返）
  ③ subproject 欄候選（最細 level）+ 有冇對應 code 欄 + sample
  ④ wynn: OPCE 家族（OPCE0NN 全部）名 + Σ → 決定邊啲 roll 去 OPCE005
  ④ mgm: 項目CAPEX-5/6（同其他 CAPEX-N）名 → 揾 golden DICJ

Run:  python scripts/inspect_project_hierarchy.py            # all 6
出 results/inspect_project_hierarchy.txt  ← paste 返嚟
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
SUBCANDS = ["Sub project", "SubProject_Name", "Subproject_Name", "Subproject", "subproject",
            "项目名称中文", "项目英文名称", "Initiative Name", "Contents Name", "Subproject_name",
            "Project/Sub-Project Name", "子項目", "Sub-Project Serial Number", "Project Serial Number"]


def _alias(s):
    sl = str(s).strip().lower()
    for kws, a in NAME2ALIAS:
        if any(k.lower() in sl for k in kws): return a
    return None


def _ndicj(v):
    v = str(v).strip()
    m = re.match(r"^(博彩項目|項目)0*(\d.*)$", v)
    return m.group(1) + m.group(2) if m else v


def _nm(s):
    """name normalize for match: 去空格/標點/全形, lower."""
    return re.sub(r"[\s\(\)（）\[\]【】,，.。:：;；\-—_/、、'\"`*]+", "", str(s).strip().lower())


def _col(df, cands):
    low = {str(c).strip().lower(): c for c in df.columns}
    for c in cands:
        if not c: continue
        if c in df.columns: return c
        if str(c).strip().lower() in low: return low[str(c).strip().lower()]
    return None


def _amtcol(df, cfg):
    a = (cfg.get("columns") or {}).get("amount", "")
    return _col(df, [a, "Debit minus Credit", "MOP Amt", "Amount - Amended", "amount_mop", "amount",
                     "Val/COArea Crcy", "Entry Voucher Amount/ Expense Amount ", "adjusted_amount"])


def main():
    want = [a.lower() for a in sys.argv[1:]] or list(ENT)
    L = ["# inspect_project_hierarchy — project名=golden DICJ名 / subproject=最細+code 準備"]
    gpath = next((p for p in GCAND if p.exists()), None)
    gmap = {}      # {alias: {dicj_norm: 名}}
    gnameset = {}  # {alias: {名_norm: 原名}}  畀 name-match check 用
    if gpath:
        g = pd.read_excel(gpath, sheet_name=TAB, dtype=str)
        g.columns = [str(c).strip() for c in g.columns]
        g["_a"] = g[next(c for c in g.columns if "承批" in c)].map(_alias)
        g["_d"] = g[next(c for c in g.columns if c.strip() in ("DICJ Code", "DICJ"))].astype(str).str.strip().map(_ndicj)
        ncol = next((c for c in g.columns if "項目名稱" in str(c) or "项目名称" in str(c)), None)
        acol = next((c for c in g.columns if c.strip() == "Amount"), None)
        g["_n"] = g[ncol].astype(str).str.strip() if ncol else ""
        g["_amt"] = pd.to_numeric(g[acol].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0.0) if acol else 0.0
        gg = g[g["_n"].ne("") & g["_n"].ne("nan")]
        for a in gg["_a"].dropna().unique():
            sub = gg[gg["_a"] == a]
            idx = sub.groupby("_d")["_amt"].apply(lambda s: s.abs().idxmax())
            gmap[a] = dict(zip(sub.loc[idx, "_d"], sub.loc[idx, "_n"]))
            gnameset[a] = {_nm(n): n for n in sub["_n"].unique()}
        L.append(f"   golden 名 map: " + ", ".join(f"{a}={len(m)}" for a, m in gmap.items()))
    else:
        L.append("   !! golden 揾唔到 → 冇名 map")

    for alias in want:
        com = ENT.get(alias)
        if not com: continue
        L.append(f"\n{'='*72}\n## {alias}")
        p = ROOT / "data" / alias / "output" / f"{com}_kpi_report.parquet"
        if not p.exists():
            L.append("   (no parquet)"); continue
        cfg = yaml.safe_load((ROOT / "conf" / com / "parameters.yml").read_text(encoding="utf-8"))
        df = pd.read_parquet(p)
        amt = pd.to_numeric(df[_amtcol(df, cfg)], errors="coerce").fillna(0.0) / 1e4
        dd = df["dicj_code"].astype(str).fillna("").str.strip().map(_ndicj) if "dicj_code" in df.columns else pd.Series("", index=df.index)
        pcol = _col(df, [(cfg.get("columns") or {}).get("project", ""), "project", "Project Name", "Project name - Amended"])
        scol = _col(df, SUBCANDS)
        proj = df[pcol].astype(str).fillna("") if pcol else pd.Series("", index=df.index)
        subp = df[scol].astype(str).fillna("") if scol else pd.Series("", index=df.index)
        gm = gmap.get(alias, {})

        # ① 覆蓋
        mapped = dd.map(lambda d: d in gm)
        blank = dd.eq("") | dd.eq("nan")
        L.append(f"   ① dicj→golden名: mapped {int(mapped.sum()):,}行 | unmapped(有dicj冇名) {int((~mapped & ~blank).sum()):,}行 Σ={amt[~mapped & ~blank].sum():,.0f}萬 | blank dicj {int(blank.sum()):,}行 Σ={amt[blank].sum():,.0f}萬")

        # ② unmapped / blank → 名
        bad = (~mapped) | blank
        if bad.any():
            L.append(f"   ② 冇 golden 名嘅 dicj (頭12，按Σ)：[dicj | project | subproject]")
            t = pd.DataFrame({"d": dd[bad], "p": proj[bad].str[:26], "s": subp[bad].str[:26], "a": amt[bad]})
            for (d, pp, ss), a in t.groupby(["d", "p", "s"])["a"].sum().abs().sort_values(ascending=False).head(12).items():
                L.append(f"        {str(d)[:14]:<14} | {pp:<26} | {ss}")

        # ⑤ name-match: 我哋 project / subproject 名 對唔對到 golden 項目名稱 (答 user 個問題)
        gnorms = set(gnameset.get(alias, {}).keys())
        if gnorms:
            L.append(f"   ⑤ 我哋 input 名 同 golden 項目名稱 對唔對上：")
            pairs = [(proj, "project名")] + ([(subp, "subproject名")] if scol else [])
            for series, label in pairs:
                nm = series[series.astype(str).str.strip().ne("")].map(_nm).drop_duplicates()
                if not len(nm): continue
                exact = nm.isin(gnorms)
                subm = nm.map(lambda x: x in gnorms or any(x and (x in g or g in x) for g in gnorms))
                L.append(f"      {label}: distinct {len(nm):,} | exact-match golden {int(exact.sum())} ({exact.mean()*100:.0f}%) | substr {int(subm.sum())} ({subm.mean()*100:.0f}%)")
                bad = set(nm[~subm.values].tolist())
                orig = series[series.map(_nm).isin(bad) & series.astype(str).str.strip().ne("")].drop_duplicates()
                if len(orig):
                    L.append("          對唔到 sample: " + " / ".join(str(x)[:26] for x in orig.head(5)))

        # ③ subproject 欄
        L.append(f"   ③ project欄='{pcol}'  subproject欄='{scol}'")
        if scol:
            L.append(f"      subproject sample: " + " / ".join(subp[subp.ne('')].drop_duplicates().head(4).tolist())[:120])
        # code 欄候選 (似 code 短值)
        code_cands = [c for c in df.columns if any(k in str(c).lower() for k in ("sub", "code", "序號", "項目編號", "wbs", "job"))][:8]
        L.append(f"      可能 subproject_code 欄: {code_cands}")

        # ④ wynn OPCE 家族 / mgm CAPEX
        if alias == "wynn":
            fam = dd[dd.str.match(r"^OPCE", na=False)]
            if len(fam):
                L.append(f"   ④ OPCE 家族 (code | Σ萬 | 1個project名)：")
                for d, a in amt[dd.str.match(r'^OPCE', na=False)].groupby(dd[dd.str.match(r'^OPCE', na=False)]).sum().abs().sort_values(ascending=False).head(20).items():
                    nm = proj[dd == d].iloc[0][:30] if (dd == d).any() else ""
                    L.append(f"        {str(d)[:12]:<12} {a:>9,.0f}萬  {nm}")
        if alias == "mgm":
            cap = dd[dd.str.contains("CAPEX", na=False)]
            if len(cap):
                L.append(f"   ④ mgm CAPEX-N (code | Σ萬 | project名)：")
                for d in sorted(cap.unique()):
                    nm = proj[dd == d].iloc[0][:34] if (dd == d).any() else ""
                    L.append(f"        {str(d)[:14]:<14} {amt[dd==d].sum():>9,.0f}萬  {nm}")
    out = ROOT / "results" / "inspect_project_hierarchy.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
