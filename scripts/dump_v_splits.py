r"""Dump V 拆分類嘅明細，驗 keyword heuristic 拆得啱唔啱 + 睇有冇更好基準（項目組V finer ref）。
拆分 V：娛樂表演/美食活動/路演/社區活化/邀請外國客人/邀請旅行社/拍攝視頻。

對每個原 vertical_id，列出我 heuristic 拆出嘅新 label 分佈 + 各自 top 項目組V / account（睇基準對唔對）。

Run:  python scripts/dump_v_splits.py
出 results/dump_v_splits.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
SPLITS = ["V_CONCERT", "V_FOOD_EVENT", "V_OVERSEAS_ROADSHOW", "V_COMMUNITY",
          "V_INVITE_GUEST", "V_INVITE_AGENCY", "V_PROMO_VIDEO"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# dump_v_splits — 拆分 V 嘅新 label 分佈 + 項目組V / account 基準（驗 heuristic）"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["amt"] = pd.to_numeric(df["amount_mop"], errors="coerce").fillna(0.0) / 1e4
    vid = _s(df["vertical_id"]); vlab = _s(df["vertical_label"])
    ptv = _s(df["項目組V"]) if "項目組V" in df.columns else pd.Series("", index=df.index)
    acc = _s(df["account_desc"]) if "account_desc" in df.columns else pd.Series("", index=df.index)
    for v in SPLITS:
        m = vid == v
        if not m.any():
            continue
        tot = df.loc[m, "amt"].sum()
        L.append(f"\n{'='*60}\n── {v}  Σ={tot:,.0f}萬 → 拆成: ──")
        sub = df[m]
        for nl, g in sorted(sub.groupby(vlab[m]), key=lambda kv: -kv[1]['amt'].sum()):
            L.append(f"   「{nl}」 {g['amt'].sum():,.0f}萬 ({len(g)}行)")
            pv = g.groupby(ptv[g.index])["amt"].sum().abs().sort_values(ascending=False).head(4)
            L.append("       項目組V: " + " | ".join(f"{str(k)[:16] or '(空)'}={x:,.0f}" for k, x in pv.items()))
            aa = g.groupby(acc[g.index])["amt"].sum().abs().sort_values(ascending=False).head(4)
            L.append("       account: " + " | ".join(f"{str(k)[:16]}={x:,.0f}" for k, x in aa.items()))
    _w(L)


def _w(L):
    out = ROOT / "results" / "dump_v_splits.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
