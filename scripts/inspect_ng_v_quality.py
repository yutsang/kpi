r"""inspect_ng_v_quality —— 診斷 V 廣播 + project 名 NG-blind 回填 兩個問題。

問題1：非該-V 嘅嘢被廣播/enforce 入某 V（如水上樂園/演唱會入「路演」）。
問題2：項目名 ← golden dicj-idxmax（NG-blind）→ 一個 dicj 跨多 NG 時冚上最大項目名。

三段：
  A. dicj_code 跨多 NG / 多項目名（量化 NG-blind 名回填污染）
  B. 指定 NG（default NG1 吸引外國客源）每 entity×project：dicj / vertical_label / 項目組V / v_調整 / 金額
     → 睇邊啲係廣播(v_調整=填補/重設)、項目組V 本身講咩、邊啲 project 名似掛錯
  C. 指定 NG×V（default 路演）by v_調整 flag —— 幾多係原始真V vs 廣播

Run（Windows）:  python scripts\inspect_ng_v_quality.py [NG] [V關鍵字]
                 e.g. python scripts\inspect_ng_v_quality.py NG1 路演
出 results\inspect_ng_v_quality.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "inspect_ng_v_quality.txt"
NG_TGT = sys.argv[1] if len(sys.argv) > 1 else "NG1"
V_TGT  = sys.argv[2] if len(sys.argv) > 2 else "路演"


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = [f"# inspect_ng_v_quality —— V廣播 + project名NG-blind回填（目標 {NG_TGT} / V含「{V_TGT}」；萬）"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["amt"] = pd.to_numeric(df.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    for c, n in [("ng_code", "ng"), ("ng_label", "nglab"), ("vertical_label", "vl"), ("vertical_id", "vid"),
                 ("project", "proj"), ("dicj code", "dicj"), ("項目組V", "ptv"), ("v_調整", "vadj"),
                 ("entity", "e")]:
        df[n] = _s(df.get(c, pd.Series("", index=df.index)))

    # ── A. dicj 跨多 NG / 多 project 名 ───────────────────────────────────────
    L.append("\n" + "=" * 80)
    L.append("## A. dicj_code 跨多 NG / 多 project 名（NG-blind 名回填污染量化）")
    by_dicj = df[df["dicj"].ne("")].groupby(["e", "dicj"]).agg(
        n_ng=("ng", "nunique"), n_proj=("proj", "nunique"), amt=("amt", "sum")).reset_index()
    multi = by_dicj[by_dicj["n_ng"] > 1]
    L.append(f"   dicj 總數={len(by_dicj):,} ｜ 跨>1個NG嘅 dicj={len(multi):,}（= 名回填可能污染）")
    L.append("   ── 跨最多 NG 嘅 dicj top 12 ──")
    for _, r in multi.sort_values("n_ng", ascending=False).head(12).iterrows():
        sub = df[(df.e.eq(r["e"])) & (df.dicj.eq(r["dicj"]))]
        ngs = ",".join(sorted(sub["ng"].unique())[:6])
        pj = sub.groupby("proj")["amt"].sum().sort_values(key=lambda s: s.abs(), ascending=False)
        pjs = " / ".join(f"{p[:14]}({a:,.0f})" for p, a in pj.head(3).items())
        L.append(f"     {r['e']:<7} dicj={r['dicj'][:14]:<14} {int(r['n_ng'])}NG[{ngs}] 名:{pjs}")

    # ── B. 目標 NG 每 entity×project ──────────────────────────────────────────
    g = df[df["ng"].eq(NG_TGT)]
    L.append("\n" + "=" * 80)
    L.append(f"## B. {NG_TGT}（{g['nglab'].iloc[0] if len(g) else ''}）每 entity×project：vl / 項目組V / v_調整 / dicj")
    L.append(f"   {'ent':<7}{'金額':>9}  {'vl(我V)':<12}{'項目組V':<14}{'v_調整':<6} project / dicj")
    gg = g.groupby(["e", "proj", "vl", "ptv", "vadj", "dicj"])["amt"].sum().reset_index()
    gg = gg.sort_values("amt", key=lambda s: s.abs(), ascending=False)
    for _, r in gg.head(40).iterrows():
        L.append(f"   {r['e']:<7}{r['amt']:>9,.0f}  {(r['vl'] or '(空)')[:12]:<12}{(r['ptv'] or '(空)')[:14]:<14}{(r['vadj'] or '原始'):<6} {r['proj'][:24]} / {r['dicj'][:12]}")

    # ── C. 目標 NG×V by v_調整 flag ───────────────────────────────────────────
    gv = g[g["vl"].str.contains(V_TGT, na=False)]
    L.append("\n" + "=" * 80)
    L.append(f"## C. {NG_TGT} × vl含「{V_TGT}」 by v_調整（廣播 vs 原始真V）：{len(gv):,}行 Σ={gv['amt'].sum():,.0f}萬")
    for flag, sub in gv.groupby(gv["vadj"].replace("", "原始")):
        L.append(f"   v_調整={flag:<6} {len(sub):,}行  Σ={sub['amt'].sum():,.0f}萬")
    L.append("   ── 當中 項目組V top（睇佢哋本身應該係咩 V）──")
    for ptv, amt in gv.groupby(gv["ptv"].replace("", "(項目組V空)"))["amt"].sum().sort_values(key=lambda s: s.abs(), ascending=False).head(12).items():
        L.append(f"     {ptv[:30]:<30}{amt:>10,.0f}萬")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
