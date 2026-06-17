r"""Diagnose 4 data-quality issues in tableau_combined_25.csv (user 2026-06-17):
  ① 冇 NG code            — ng_code 空 / 唔係 NG\d+
  ② subproject 名含 code + 重疊 — 同一 dicj code 出現 dicj層碼 + 細碼兩種（galaxy A1.1 vs A001）
  ③ 有項目名但冇數字       — (entity, subproject code) 有名但 Σ amount = 0
  ④ NG×V scatter         — 每個 (entity, ng) 散落幾多 V label，top V 佔幾多 %

Run:  python scripts/diag_data_quality.py
出 results/diag_data_quality.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys, re
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
NGRE = re.compile(r"^NG\d+$")


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# diag_data_quality — 4 問題量化 (no-NG / subproject碼重疊 / 有名冇數 / NG×V scatter)"]
    if not CSV.exists():
        L.append("!! tableau_combined_25.csv 揾唔到 — 先跑 prep_tableau"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["amt"] = pd.to_numeric(df["amount_mop"], errors="coerce").fillna(0.0)
    ent = df["entity"].astype(str)
    ng = _s(df["ng_code"]); ngl = _s(df["ng_label"])
    dcode = _s(df["dicj code"]); scode = _s(df["subproject code"])
    sp = _s(df["subproject"]); vl = _s(df["vertical_label"])

    # ── ① 冇 NG code ──
    bad = ~ng.str.match(NGRE)
    L.append(f"\n① 冇 NG code（空/非 NG\\d+）: {int(bad.sum()):,} 行  Σ={df.loc[bad,'amt'].sum()/1e6:,.1f}M")
    for e, g in df[bad].groupby(ent):
        vals = [v for v in g['ng_code'].astype(str).str.strip().unique()][:6]
        L.append(f"     {e:<7} {len(g):>7,} 行 Σ={g['amt'].sum()/1e6:>9,.1f}M  ng_code值={vals}")
        top = g.groupby(_s(g['dicj code']))['amt'].sum().abs().sort_values(ascending=False).head(3)
        L.append(f"        top dicj: {[f'{k}={v/1e6:.0f}M' for k,v in top.items()]}")

    # ── ② subproject code 重疊（dicj層碼 leaked 入 subproject code）──
    # 對每個 (entity, dicj code)：啲 subproject code 入面有冇等於 dicj code 自己（=dicj層）+ 又有其他細碼
    L.append("\n② subproject code 重疊（同一 dicj 有 dicj層碼 + 細碼兩種 = 重複）:")
    _nh = sum(1 for c, n in zip(scode.tolist(), sp.tolist()) if c and c in n)
    L.append(f"   subproject 名包含 subproject code 嘅行: {_nh:,}")
    g = pd.DataFrame({"e": ent, "d": dcode, "s": scode, "a": df["amt"]})
    g = g[(g["d"] != "") & (g["s"] != "")]
    ov_rows = []
    for (e, d), sub in g.groupby(["e", "d"]):
        codes = set(sub["s"].unique())
        if d in codes and len(codes) > 1:   # dicj code 本身出現 + 仲有其他碼 = 重疊
            ov_rows.append((e, d, len(codes), sub.loc[sub["s"] == d, "a"].sum() / 1e6,
                            sorted(codes - {d})[:4]))
    L.append(f"   重疊 dicj code 數: {len(ov_rows)}（dicj 碼自己 leaked 做 subproject code）")
    for e, d, nc, dicjamt, others in sorted(ov_rows, key=lambda t: -abs(t[3]))[:12]:
        L.append(f"     {e:<7} {d:<10} 有 {nc} 種碼  dicj層那批Σ={dicjamt:,.0f}M  其他碼={others}")

    # ── ③ 有項目名但冇數字（subproject 有名，Σ amount ≈ 0）──
    L.append("\n③ 有 subproject 名但 Σ amount = 0:")
    grp = pd.DataFrame({"e": ent, "s": scode, "sp": sp, "a": df["amt"]})
    grp = grp[grp["s"] != ""]
    agg = grp.groupby(["e", "s"]).agg(amt=("a", "sum"), n=("a", "size"),
                                       name=("sp", "first")).reset_index()
    zero = agg[agg["amt"].abs() < 1e-6]
    L.append(f"   {len(zero):,} 個 (entity, subproject code) 有名但 Σ=0（共 {int(zero['n'].sum()):,} 行）")
    for e, sub in zero.groupby("e"):
        L.append(f"     {e:<7} {len(sub):>4} 個  e.g. {[f'{r.s}:{str(r.name)[:18]}' for r in sub.head(4).itertuples()]}")

    # ── ④ NG×V scatter ──
    L.append("\n④ NG×V scatter（每個 entity×NG 散落幾多 V，top V 佔 %）:")
    q = pd.DataFrame({"e": ent, "ng": ng, "ngl": ngl, "v": vl, "a": df["amt"].abs()})
    q = q[q["ng"].str.match(NGRE) & (q["v"] != "")]
    for (e, ngc), sub in q.groupby(["e", "ng"]):
        byv = sub.groupby("v")["a"].sum().sort_values(ascending=False)
        tot = byv.sum()
        if tot <= 0:
            continue
        topshare = byv.iloc[0] / tot * 100
        if len(byv) >= 4 and topshare < 75:   # 散 + top 唔夠主導 → flag
            L.append(f"     {e:<7} {ngc:<5} {str(sub['ngl'].iloc[0])[:6]:<6} {len(byv)}個V top={byv.index[0][:8]}({topshare:.0f}%)  "
                     f"次={[f'{k[:6]}={v/tot*100:.0f}%' for k,v in byv.iloc[1:4].items()]}")
    _w(L)


def _one(x):
    return str(x).strip()


def _w(L):
    out = ROOT / "results" / "diag_data_quality.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
