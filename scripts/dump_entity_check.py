r"""一家 entity 嘅最終 deliverable 驗證 dump（default sjm）—— 睇實數：
  ① bucket × 調整前/調整/調整後 總額（tie 用）
  ② 新 V label × 調整後（投資方向 V 分佈）
  ③ H label × 調整後（橫向分佈）
  ④ NG × 新V cross-tab（top，睇合唔合理）

Run:  python scripts/dump_entity_check.py            # sjm
      python scripts/dump_entity_check.py vml
出 results/dump_entity_check.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
ENT = sys.argv[1] if len(sys.argv) > 1 else "sjm"


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def _num(s):
    return pd.to_numeric(s, errors="coerce").fillna(0.0)


def main():
    L = [f"# dump_entity_check — {ENT} 最終 deliverable 驗證（萬元）"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df = df[df["entity"].astype(str) == ENT]
    if not len(df):
        L.append(f"!! {ENT} 冇 row"); _w(L); return
    pre = _num(df.get("調整前_萬", 0)); adj = _num(df.get("調整_萬", 0)); post = _num(df.get("調整後_萬", 0))
    b = _s(df["year_bucket"]); v = _s(df["vertical_label"]); h = _s(df["horizontal_label"]); ng = _s(df["ng_label"])

    L.append(f"\n① bucket × 調整前/調整/調整後（萬）：")
    g = pd.DataFrame({"b": b, "pre": pre, "adj": adj, "post": post}).groupby("b").sum()
    for i, r in g.iterrows():
        L.append(f"     {i:<8} 調整前={r.pre:>12,.0f}  調整={r.adj:>11,.0f}  調整後={r.post:>12,.0f}")
    L.append(f"     {'總計':<8} 調整前={pre.sum():>12,.0f}  調整={adj.sum():>11,.0f}  調整後={post.sum():>12,.0f}")

    L.append(f"\n② 新 V label × 調整後（萬）：")
    gv = pd.DataFrame({"v": v, "post": post}).groupby("v")["post"].sum().sort_values(ascending=False)
    for k, x in gv.items():
        L.append(f"     {str(k)[:24]:<24} {x:>12,.0f}")

    L.append(f"\n③ H label × 調整後（萬）：")
    gh = pd.DataFrame({"h": h, "post": post}).groupby("h")["post"].sum().sort_values(ascending=False)
    for k, x in gh.items():
        L.append(f"     {str(k)[:18]:<18} {x:>12,.0f}")

    L.append(f"\n④ NG × 新V（調整後萬，每 NG top 5）：")
    q = pd.DataFrame({"ng": ng, "v": v, "post": post})
    for n, sub in q.groupby("ng"):
        top = sub.groupby("v")["post"].sum().sort_values(ascending=False).head(5)
        L.append(f"   {n}: " + " | ".join(f"{str(k)[:10]}={x:,.0f}" for k, x in top.items()))
    _w(L)


def _w(L):
    out = ROOT / "results" / "dump_entity_check.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
