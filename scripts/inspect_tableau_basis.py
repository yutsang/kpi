r"""對清楚 Tableau 到底 sum 緊邊個金額欄 —— comp/staff per entity，
amount_mop vs 調整前_萬 vs 調整後_萬，exact bucket vs 25-prefix。
同 user Tableau 個表比，搵返 Tableau 用緊邊個 field + 邊個 aggregation。

user 表（萬）comp 我哋: galaxy 58,625/89,031/92,524; staff: galaxy 13,868/18,614/27,786
Run（Windows）:  python scripts\inspect_tableau_basis.py
出 results\inspect_tableau_basis.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
ENTS = ["galaxy", "wynn", "vml", "melco", "mgm", "sjm"]
COMP_H = {"H_HOTEL_ROOM", "H_VENUE", "H_FNB", "H_COMP_TICKET", "H_COMP_OTHER"}
L = ["# inspect_tableau_basis — comp/staff × {amount_mop, 調整前_萬, 調整後_萬} × bucket（萬）"]


def main():
    if not CSV.exists():
        print("!! csv 揾唔到"); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    amt = pd.to_numeric(df.get("amount_mop", 0), errors="coerce").fillna(0.0) / 1e4   # 萬
    pre = pd.to_numeric(df.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    post = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    df["amt"] = amt; df["pre"] = pre; df["post"] = post
    yb = df["year_bucket"].astype(str)
    df["yb"] = yb; df["y2"] = yb.str[:2]
    df["ent"] = df["entity"].astype(str)
    hid = df.get("horizontal_id", pd.Series("", index=df.index)).astype(str).str.strip()
    co = df.get("final_capex_opex", pd.Series("", index=df.index)).astype(str).str.strip()
    df["isComp"] = hid.isin(COMP_H)
    df["isStaffO"] = hid.eq("H_LABOR") & ~co.eq("Capex")   # opex labor
    df["isStaffAll"] = hid.eq("H_LABOR")                    # incl capex

    def cell(ent, mask, bk, pref=False):
        if pref:
            m = (df.ent == ent) & mask & (df.y2 == bk)
        else:
            m = (df.ent == ent) & mask & (df.yb == bk)
        return df.loc[m, "amt"].sum(), df.loc[m, "pre"].sum(), df.loc[m, "post"].sum()

    for kind, mask in [("COMP", df.isComp), ("STAFF_opex", df.isStaffO), ("STAFF_all(incl capex)", df.isStaffAll)]:
        L.append(f"\n{'='*78}\n## {kind}  （每格 = amt_mop | 調整前 | 調整後）")
        L.append(f"   {'ent':<8}{'23(exact)':>26}{'24(exact)':>26}{'25(prefix)':>26}")
        for ent in ENTS:
            c23 = cell(ent, mask, "23")
            c24 = cell(ent, mask, "24")
            c25 = cell(ent, mask, "25", pref=True)
            def f(t): return f"{t[0]:>7,.0f}|{t[1]:>7,.0f}|{t[2]:>7,.0f}"
            L.append(f"   {ent:<8}{f(c23):>26}{f(c24):>26}{f(c25):>26}")
    out = ROOT / "results" / "inspect_tableau_basis.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
