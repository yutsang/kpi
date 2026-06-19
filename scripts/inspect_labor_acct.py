r"""通用：一家 entity 嘅 opex 人工成本 by account_code + account_desc per bucket，
喺 RECON 金額 basis（23/24=調整後，25=調整前）—— 砌 per-year staff 規律對 HQ ±500。

Run（Windows）:  python scripts\inspect_labor_acct.py vml
                 python scripts\inspect_labor_acct.py galaxy
                 python scripts\inspect_labor_acct.py melco mgm
出 results\inspect_labor_acct_<ent>.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
ENTS = sys.argv[1:] or ["vml"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def run(ent, df):
    L = [f"# inspect_labor_acct — {ent} opex 人工 by account（萬，recon basis：23/24=調整後,25=調整前）"]
    e = df[df["entity"].astype(str) == ent].copy()
    post = pd.to_numeric(e.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    pre = pd.to_numeric(e.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    yb = e["year_bucket"].astype(str)
    y2 = yb.str[:2]
    e["m"] = post.where(yb.isin(["23", "24"]), pre)   # exact-bucket recon basis（staff 用 exact）
    e["bk"] = yb
    for c in ("horizontal_id", "final_capex_opex", "account_code", "account_desc"):
        e[c] = _s(e[c]) if c in e.columns else ""
    e["code"] = e["account_code"].str.replace(r"\.0$", "", regex=True)
    opex = ~e["final_capex_opex"].eq("Capex")
    lab = e[(e.horizontal_id == "H_LABOR") & opex]
    for yb_v in ("23", "24", "25"):
        sub = lab[lab.bk == yb_v]
        if not len(sub):
            continue
        L.append(f"\n{'='*60}\n## {ent} bucket {yb_v}  人工 Σ={sub['m'].sum():,.0f}萬（{len(sub):,}行）")
        g = sub.groupby(["code", "account_desc"])["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
        for _, r in g.head(15).iterrows():
            L.append(f"   {r['code']:<11}{r['account_desc'][:30]:<30} {r['m']:>9,.0f}萬")
    out = ROOT / "results" / f"inspect_labor_acct_{ent}.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟\n")


def main():
    if not CSV.exists():
        print("!! csv 揾唔到"); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    for ent in ENTS:
        run(ent, df)


if __name__ == "__main__":
    main()
