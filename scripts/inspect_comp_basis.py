r"""決定性測試：comp/staff 每 component 對 golden，比較 調整後_萬 vs 調整前_萬 邊個 match。
假設：golden = 調整前（gross）；大 adjustment 家（galaxy/wynn）調整後(net)會 under。
若 調整前 對得返 → comp basis 改用 調整前_萬（零 row move、零注入）。

Run（Windows）:  python scripts\inspect_comp_basis.py
出 results\inspect_comp_basis.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
GOLD = ROOT / "internal_resource_golden.xlsx"
ENTS = ["galaxy", "wynn", "vml", "melco", "mgm", "sjm"]
# component → (golden col, our horizontal_id)
COMP = [("人工", "internal_staff_cost", "H_LABOR"), ("客房", "客房支出", "H_HOTEL_ROOM"),
        ("會場", "會場支出", "H_VENUE"), ("食飲", "食品飲料支出", "H_FNB"),
        ("贈票", "贈票支出", "H_COMP_TICKET"), ("comp其他", "comp_其他", "H_COMP_OTHER")]
L: list[str] = ["# inspect_comp_basis —— comp/staff 每 component：golden vs 調整後 vs 調整前（萬，23/24）",
                "# Δ後=調整後−golden  Δ前=調整前−golden  → 邊個細邊個啱"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    if not CSV.exists() or not GOLD.exists():
        print("!! csv/golden 揾唔到"); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["post"] = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    df["pre"] = pd.to_numeric(df.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    df["yb"] = df["year_bucket"].astype(str)
    df["ent"] = df["entity"].astype(str)
    hid = _s(df.get("horizontal_id", pd.Series("", index=df.index)))
    g = pd.read_excel(GOLD); g.columns = [str(c).strip() for c in g.columns]
    g["entity"] = g["entity"].astype(str).str.strip()
    g["year"] = g["year"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    for _, c, _ in COMP:
        g[c] = pd.to_numeric(g.get(c, 0), errors="coerce").fillna(0.0)
    for ent in ENTS:
        for yr in ("23", "24"):
            gy = g[(g.entity == ent) & (g.year == yr)]
            if not len(gy):
                continue
            L.append(f"\n{'='*72}\n## {ent} {yr}")
            L.append(f"   {'comp':<8}{'golden':>10}{'調整後':>10}{'Δ後':>9}{'調整前':>10}{'Δ前':>9}  贏")
            m_ent = (df.ent == ent) & (df.yb == yr)
            for lab, gcol, h in COMP:
                gv = gy[gcol].sum()
                mh = m_ent & hid.eq(h)
                po = df.loc[mh, "post"].sum()
                pr = df.loc[mh, "pre"].sum()
                if lab == "人工" and gv == 0:   # 24 staff 喺 per-NG 檔係 hole → skip
                    continue
                win = "前" if abs(pr - gv) < abs(po - gv) else "後"
                L.append(f"   {lab:<8}{gv:>10,.0f}{po:>10,.0f}{po-gv:>9,.0f}{pr:>10,.0f}{pr-gv:>9,.0f}   {win}")
    out = ROOT / "results" / "inspect_comp_basis.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
