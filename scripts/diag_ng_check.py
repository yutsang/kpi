r"""Check NG 有冇填錯 —— 同一 dicj_code（= golden 一個項目）理應同一個 NG。
揾兩樣：
  ① 一個 dicj_code 出現多個 NG（內部矛盾，多數填錯）—— by 金額
  ② 每 dicj_code 嘅主 NG + golden 名（project 欄）—— 畀你眼睇 NG 啱唔啱（名係體育→NG應4 等）

Run:  python scripts/diag_ng_check.py
出 results/diag_ng_check.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# diag_ng_check — 同一 dicj_code 多個 NG（填錯訊號）+ dicj→主NG→golden名（眼睇）"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["amt"] = pd.to_numeric(df["amount_mop"], errors="coerce").fillna(0.0)
    ent = df["entity"].astype(str)
    dc = _s(df["dicj code"]) if "dicj code" in df.columns else pd.Series("", index=df.index)
    proj = _s(df["project"]) if "project" in df.columns else pd.Series("", index=df.index)
    ngc = _s(df["ng_code"]); ngl = _s(df["ng_label"])
    q = pd.DataFrame({"e": ent, "d": dc, "proj": proj, "ngc": ngc, "ngl": ngl, "a": df["amt"]})
    q = q[q["d"] != ""]

    # ① 一個 dicj_code 多個 NG
    L.append("\n① 同一 dicj_code 出現多個 NG（內部矛盾，多數填錯）by 金額：")
    multi = []
    for (e, d), sub in q.groupby(["e", "d"]):
        ngs = sub.groupby("ngc").agg(a=("a", "sum"), ngl=("ngl", "first"))
        if len(ngs) > 1:
            nm = sub["proj"].mode().iloc[0] if not sub["proj"].mode().empty else ""
            multi.append((e, d, nm, abs(sub["a"].sum()) / 1e4,
                          [(i, r.ngl, r.a / 1e4) for i, r in ngs.sort_values("a", ascending=False).iterrows()]))
    multi.sort(key=lambda t: -t[3])
    L.append(f"   {len(multi)} 個 dicj_code 有 >1 NG")
    for e, d, nm, tot, ngs in multi[:25]:
        ss = " / ".join(f"{i}{lab[:4]}={a:,.0f}萬" for i, lab, a in ngs)
        L.append(f"     {e:<7} {d:<10} {str(nm)[:20]:<20} → {ss}")

    # ② per entity：dicj → 主 NG + golden 名（眼睇 NG 啱唔啱），按金額 top
    L.append("\n② dicj → 主NG + golden名（眼睇 NG 合唔合理，每家 top 15）：")
    for e, sub in q.groupby("e"):
        L.append(f"\n   ── {e} ──")
        gg = sub.groupby("d").agg(a=("a", "sum"),
                                  ngl=("ngl", lambda s: s.mode().iloc[0] if not s.mode().empty else ""),
                                  proj=("proj", lambda s: s.mode().iloc[0] if not s.mode().empty else "")).reset_index()
        gg = gg.reindex(gg["a"].abs().sort_values(ascending=False).index)
        for r in gg.head(15).itertuples():
            L.append(f"     {r.d:<10} [{str(r.ngl)[:5]:<5}] {str(r.proj)[:32]:<32} {r.a/1e4:>10,.0f}萬")
    _w(L)


def _w(L):
    out = ROOT / "results" / "diag_ng_check.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
