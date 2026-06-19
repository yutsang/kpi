r"""揾每家 CAPEX 入面似 staff 但 tag 咗 建設/設施 嘅行（capex staff 攤入建設）。
信號：account_code prefix（各家 staff 碼）/ 項目組H 含 staff / account_desc 似 staff。
目標：搵返 capex staff 補入人工（user Table 1 capex+opex 調整前：galaxy24=34,926 等）。

Run（Windows）:  python scripts\inspect_capex_acct.py galaxy vml melco mgm wynn sjm
出 results\inspect_capex_acct_<ent>.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
ENTS = sys.argv[1:] or ["galaxy", "vml", "melco", "mgm", "wynn", "sjm"]
# 各家 staff account_code prefix（由 opex 人工 dump 得知）
STAFF_CODE = {
    "galaxy": ("8000", "8020", "8103", "7926", "8221"),
    "vml": ("70110", "70111", "70113", "70130", "70140", "70171", "73", "76100", "80180"),
    "melco": ("611", "841"),
    "mgm": ("521", "522", "588050"),
    "wynn": (), "sjm": (),
}
SDESC = ["Salar", "salar", "SALAR", "Payroll", "payroll", "Casual", "casual", "Staff", "staff",
         "Wage", "wage", "Contract Labo", "CONTRACT LABO", "Bonus", "Provident", "Overtime",
         "Basic Salary", "Allowance", "Pension", "人工", "薪", "工資", "Employee Salar"]
L: list[str] = ["# inspect_capex_acct —— CAPEX 入面似 staff 但 tag 咗建設/設施（萬=調整後_萬）"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def run(ent, df):
    out = [f"# {ent} CAPEX staff scan（萬）"]
    e = df[df["entity"].astype(str) == ent].copy()
    e["post"] = pd.to_numeric(e.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    e["yb"] = e["year_bucket"].astype(str)
    for c in ("horizontal_id", "horizontal_label", "final_capex_opex", "account_code", "account_desc", "項目組H"):
        e[c] = _s(e[c]) if c in e.columns else ""
    e["code"] = e["account_code"].str.replace(r"\.0$", "", regex=True)
    cap = e["final_capex_opex"].eq("Capex")
    codes = STAFF_CODE.get(ent, ())
    sig_code = e["code"].apply(lambda c: any(c.startswith(p) for p in codes)) if codes else pd.Series(False, index=e.index)
    sig_desc = e["account_desc"].apply(lambda s: any(k in str(s) for k in SDESC))
    sig_pth = e["項目組H"].str.contains("staff|Salar|Payroll|人工", case=False, na=False)
    staff_like = cap & (sig_code | sig_desc | sig_pth)
    notlab = staff_like & e["horizontal_id"].ne("H_LABOR")
    out.append(f"  CAPEX staff-like 非人工：Σ={e.loc[notlab,'post'].sum():,.0f}萬（{int(notlab.sum()):,}行）")
    out.append("  ── by bucket ──")
    for yb in ("23", "24", "25"):
        m = notlab & e.yb.eq(yb)
        if int(m.sum()):
            out.append(f"     {yb}: {e.loc[m,'post'].sum():,.0f}萬")
    out.append("  ── top account_code（現H）──")
    g = e[notlab].groupby(["code", "account_desc", "horizontal_label"])["post"].sum().reset_index().sort_values("post", key=lambda s: s.abs(), ascending=False)
    for _, r in g.head(15).iterrows():
        out.append(f"     {r['code']:<10}{r['account_desc'][:26]:<26} 現H={r['horizontal_label'][:8]:<8} {r['post']:>8,.0f}萬")
    p = ROOT / "results" / f"inspect_capex_acct_{ent}.txt"; p.parent.mkdir(exist_ok=True)
    p.write_text("\n".join(out), encoding="utf-8")
    print("\n".join(out)); print(f"wrote {p.relative_to(ROOT)}\n")


def main():
    if not CSV.exists():
        print("!! csv 揾唔到"); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    for ent in ENTS:
        run(ent, df)


if __name__ == "__main__":
    main()
