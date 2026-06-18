r"""NG 多值清單（BUCKET-AWARE）—— 答：
  #3 邊啲 golden(dicj) 喺同一 bucket 含多個 NG（你 by-NG tie 到 = 呢啲就係 golden 真實多 NG）
  #2 每個 NG 嘅 調整前 / 調整 / 調整後（萬）—— 睇實如果某 NG 唔啱，金額影響有幾大

NG = databook 主題欄，你已 by (entity,bucket,NG) tie golden → NG 本身啱，呢個清單係參考用。

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


def _num(s):
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def main():
    L = ["# diag_ng_check — golden 含多 NG 嘅 dicj×bucket 清單 + 每 NG 調整前/調整/調整後(萬)"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    pre = _num(df["調整前_萬"]) if "調整前_萬" in df.columns else _num(df.get("amount_mop", 0)) / 1e4
    adj = _num(df["調整_萬"]) if "調整_萬" in df.columns else pd.Series(0.0, index=df.index)
    post = _num(df["調整後_萬"]) if "調整後_萬" in df.columns else pre
    q = pd.DataFrame({
        "e": df["entity"].astype(str),
        "d": _s(df["dicj code"]) if "dicj code" in df.columns else "",
        "b": _s(df["year_bucket"]) if "year_bucket" in df.columns else "",
        "proj": _s(df["project"]) if "project" in df.columns else "",
        "ngc": _s(df["ng_code"]), "ngl": _s(df["ng_label"]),
        "pre": pre, "adj": adj, "post": post,
    })
    q = q[q["d"] != ""]

    rows = []
    for (e, d, b), sub in q.groupby(["e", "d", "b"]):
        by = sub.groupby(["ngc", "ngl"]).agg(pre=("pre", "sum"), adj=("adj", "sum"), post=("post", "sum"))
        by = by[(by["post"].abs() > 0.5) | (by["pre"].abs() > 0.5)]
        if len(by) > 1:
            nm = sub["proj"].mode().iloc[0] if not sub["proj"].mode().empty else ""
            rows.append((e, d, b, nm, by["post"].abs().sum(), by.sort_values("post", ascending=False)))
    rows.sort(key=lambda t: -t[4])

    L.append(f"\n共 {len(rows)} 個 (dicj, bucket) 含 >1 NG。逐個列 調整前/調整/調整後(萬)：\n")
    for e, d, b, nm, tot, by in rows:
        L.append(f"── {e} {d} [{b}] {str(nm)[:30]} ──")
        for (ngc, ngl), r in by.iterrows():
            L.append(f"     {ngc:<5}{str(ngl)[:6]:<6}  調整前={r.pre:>11,.1f}  調整={r.adj:>10,.1f}  調整後={r.post:>11,.1f}")
    _w(L)


def _w(L):
    out = ROOT / "results" / "diag_ng_check.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
