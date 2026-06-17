r"""Check NG 填錯 —— BUCKET-AWARE（user: 一個項目跨 bucket 有唔同 NG 係 golden 真實、合理；
只有同一 bucket 內一個 dicj 出現多 NG 先係真錯）。

  ① 真錯：同一 (dicj_code, bucket) 出現 >1 NG —— by 金額，詳列每個 NG + golden 名
  ② 參考：multi-NG dicj 嘅 bucket × NG 分佈（睇係咪純跨 bucket 變化 = 合理）

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
    L = ["# diag_ng_check (BUCKET-AWARE) — 同一 (dicj, bucket) 多 NG = 真錯；跨 bucket 變化 = 合理"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["amt"] = pd.to_numeric(df["amount_mop"], errors="coerce").fillna(0.0)
    q = pd.DataFrame({
        "e": df["entity"].astype(str),
        "d": _s(df["dicj code"]) if "dicj code" in df.columns else "",
        "b": _s(df["year_bucket"]) if "year_bucket" in df.columns else "",
        "proj": _s(df["project"]) if "project" in df.columns else "",
        "ngc": _s(df["ng_code"]), "ngl": _s(df["ng_label"]), "a": df["amt"],
    })
    q = q[q["d"] != ""]

    # ① 真錯：同一 (e, d, b) 多 NG
    L.append("\n① 真錯：同一 (dicj_code, bucket) 出現 >1 NG（by 金額 top 30）：")
    rows = []
    for (e, d, b), sub in q.groupby(["e", "d", "b"]):
        ngs = sub.groupby("ngc").agg(a=("a", "sum"), ngl=("ngl", "first"))
        if len(ngs) > 1:
            nm = sub["proj"].mode().iloc[0] if not sub["proj"].mode().empty else ""
            rows.append((e, d, b, nm, abs(sub["a"].sum()) / 1e4,
                         [(i, r.ngl, r.a / 1e4) for i, r in ngs.sort_values("a", ascending=False).iterrows()]))
    rows.sort(key=lambda t: -t[4])
    L.append(f"   {len(rows)} 個 (dicj, bucket) 有 >1 NG  ← 呢啲先係真錯")
    for e, d, b, nm, tot, ngs in rows[:30]:
        ss = " / ".join(f"{i}{lab[:4]}={a:,.0f}萬" for i, lab, a in ngs)
        L.append(f"     {e:<7} {d:<9} [{b:<7}] {str(nm)[:20]:<20} → {ss}")

    # ② 參考：multi-NG dicj 嘅 bucket × NG（睇係咪純跨 bucket = 合理）
    L.append("\n② 參考：跨 bucket 有唔同 NG 嘅 dicj（每 bucket 一個 NG = 合理，唔使理）：")
    multi_d = [k for k, sub in q.groupby(["e", "d"]) if sub["ngc"].nunique() > 1]
    cross_only = 0
    shown = 0
    for (e, d), sub in q.groupby(["e", "d"]):
        if sub["ngc"].nunique() <= 1:
            continue
        per_bkt = sub.groupby("b")["ngc"].nunique()
        within = (per_bkt > 1).any()   # 有 bucket 內多 NG = 真錯（已喺①）
        if not within:
            cross_only += 1
            if shown < 12:
                nm = sub["proj"].mode().iloc[0] if not sub["proj"].mode().empty else ""
                bk = sub.groupby("b").agg(ng=("ngl", lambda s: s.mode().iloc[0] if not s.mode().empty else ""),
                                          a=("a", "sum"))
                ss = "  ".join(f"{i}:{r.ng[:4]}({r.a/1e4:,.0f}萬)" for i, r in bk.iterrows())
                L.append(f"     {e:<7} {d:<9} {str(nm)[:18]:<18} {ss}")
                shown += 1
    L.append(f"   （純跨 bucket 變化 = {cross_only} 個 dicj，合理唔使理；真錯喺①）")
    _w(L)


def _w(L):
    out = ROOT / "results" / "diag_ng_check.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
