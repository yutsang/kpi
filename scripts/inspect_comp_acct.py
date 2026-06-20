r"""一家 entity 嘅 comp（5 類）by comp_type/account_desc per bucket（recon basis：23/24=調整後,25=調整前）。
搵 comp 對唔上 HQ 嘅來源（galaxy comp 超 +10k）。

Run（Windows）:  python scripts\inspect_comp_acct.py galaxy
                 python scripts\inspect_comp_acct.py wynn mgm
出 results\inspect_comp_acct_<ent>.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
ENTS = sys.argv[1:] or ["galaxy"]
COMP_H = {"H_HOTEL_ROOM": "Comp房間", "H_VENUE": "Comp活動場地", "H_FNB": "Comp餐飲",
          "H_COMP_TICKET": "Comp贈票", "H_COMP_OTHER": "Comp其他"}


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def run(ent, df):
    L = [f"# inspect_comp_acct — {ent} comp by H/account（萬，23/24=調整後,25=調整前）"]
    e = df[df["entity"].astype(str) == ent].copy()
    post = pd.to_numeric(e.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    pre = pd.to_numeric(e.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    yb = e["year_bucket"].astype(str)
    e["m"] = post.where(yb.isin(["23", "24"]), pre)
    e["bk"] = yb
    for c in ("horizontal_id", "comp_type", "account_desc", "account_code"):
        e[c] = _s(e[c]) if c in e.columns else ""
    comp = e[e["horizontal_id"].isin(COMP_H)]
    for yb_v in ("23", "24", "25"):
        sub = comp[comp.bk == yb_v]
        if not len(sub):
            continue
        L.append(f"\n{'='*60}\n## {ent} bucket {yb_v}  comp合計 Σ={sub['m'].sum():,.0f}萬")
        for hid, lab in COMP_H.items():
            hsub = sub[sub.horizontal_id == hid]
            if abs(hsub["m"].sum()) < 1:
                continue
            L.append(f"  ── {lab} Σ={hsub['m'].sum():,.0f}萬 ──")
            g = hsub.groupby(["comp_type", "account_desc"])["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
            for _, r in g.head(6).iterrows():
                L.append(f"      ct={r['comp_type'][:14]:<14} {r['account_desc'][:26]:<26} {r['m']:>8,.0f}萬")
    p = ROOT / "results" / f"inspect_comp_acct_{ent}.txt"; p.parent.mkdir(exist_ok=True)
    p.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"wrote {p.relative_to(ROOT)}\n")


def main():
    if not CSV.exists():
        print("!! csv 揾唔到"); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    for ent in ENTS:
        run(ent, df)


if __name__ == "__main__":
    main()
