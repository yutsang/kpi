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


HQ_STAFF = {  # opex only 調整後（Table 2），年份 = recon_hq prefix key
    "galaxy": {"23": 17082, "24": 14574, "25": 23950},
    "wynn":   {"23": 6928,  "24": 8765,  "25": 10958},
    "vml":    {"23": 790,   "24": 1457,  "25": 1253},
    "melco":  {"23": 7218,  "24": 13912, "25": 27806},
    "mgm":    {"23": 6357,  "24": 6829,  "25": 21467},
    "sjm":    {"23": 0,     "24": 0,     "25": 142},
}


def run(ent, df):
    L = [f"# inspect_labor_acct — {ent} opex 人工 by account（萬，recon basis）"]
    e = df[df["entity"].astype(str) == ent].copy()
    post = pd.to_numeric(e.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    pre  = pd.to_numeric(e.get("調整前_萬",  0), errors="coerce").fillna(0.0)
    yb   = e["year_bucket"].astype(str)
    y2   = yb.str[:2]
    # recon_hq basis: 23=exact+調整後, 24-prefix=調整後, 25-prefix=調整前
    e["m_pref"] = post.where(y2.isin(["23", "24"]), pre)   # prefix basis（與 recon_hq staff 一致）
    e["y2"]  = y2
    e["bk"]  = yb
    for c in ("horizontal_id", "final_capex_opex", "account_code", "account_desc"):
        e[c] = _s(e[c]) if c in e.columns else ""
    e["code"] = e["account_code"].str.replace(r"\.0$", "", regex=True)
    opex = ~e["final_capex_opex"].eq("Capex")
    lab  = e[(e.horizontal_id == "H_LABOR") & opex]
    hq_e = HQ_STAFF.get(ent, {})
    for yr in ("23", "24", "25"):
        if yr == "23":
            sub = lab[lab.bk == "23"]
        else:
            sub = lab[lab.y2 == yr]
        if not len(sub):
            continue
        tot  = sub["m_pref"].sum()
        hq   = hq_e.get(yr, 0)
        flag = f"Δ={tot-hq:+,.0f}" if hq else ""
        L.append(f"\n{'='*62}\n## {ent} {yr}-prefix  人工 Σ={tot:,.0f}萬  HQ={hq:,}  {flag}（{len(sub):,}行）")
        g = sub.groupby(["code", "account_desc"])["m_pref"].sum().reset_index().sort_values("m_pref", key=lambda s: s.abs(), ascending=False)
        for _, r in g.head(18).iterrows():
            L.append(f"   {r['code']:<12}{r['account_desc'][:32]:<32} {r['m_pref']:>9,.0f}萬")
        # also show by exact bucket if multiple
        if yr != "23":
            L.append(f"\n   ── exact bucket breakdown ──")
            bg = sub.groupby("bk")["m_pref"].sum().reset_index().sort_values("bk")
            for _, r in bg.iterrows():
                L.append(f"      {r['bk']:<14} {r['m_pref']:>9,.0f}萬")
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
