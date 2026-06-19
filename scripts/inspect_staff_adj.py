r"""查 opex staff（人工成本）嘅「調整」係咪套到 —— HQ = opex staff 調整後，我哋對唔上(melco 7,989 vs 0)。

每家：
  §A opex H_LABOR by year_bucket：Σ調整前 / Σ調整 / Σ調整後（睇調整有冇 reduce 到）
  §B 全部「調整金額≠0」嘅行 by (調整一級, 調整二級) → Σ調整 + 落咗咩 H（睇 staff 調整係咪 tag 咗去第度）
  §C staff 調整關鍵字行（支持部門人工/內部人工/超過可計入/內部資源/payroll adj）→ H × bucket × Σ調整後

Run（Windows）:  python scripts\inspect_staff_adj.py
出 results\inspect_staff_adj.txt  ← paste 返嚟
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
L: list[str] = ["# inspect_staff_adj —— opex staff 調整有冇套到（萬）"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["pre"] = pd.to_numeric(df.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    df["adj"] = pd.to_numeric(df.get("調整_萬", 0), errors="coerce").fillna(0.0)
    df["post"] = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    df["yb"] = df["year_bucket"].astype(str)
    for c in ("entity", "horizontal_id", "horizontal_label", "final_capex_opex",
              "調整一級", "調整二級", "account_desc", "description", "科目明細", "adjust_lv1", "adjust_lv2"):
        if c not in df.columns:
            df[c] = ""
        df[c] = _s(df[c])
    lv1 = df["調整一級"].where(df["調整一級"].ne(""), df["adjust_lv1"])
    lv2 = df["調整二級"].where(df["調整二級"].ne(""), df["adjust_lv2"])
    df["blob"] = df["account_desc"] + " | " + df["description"] + " | " + df["科目明細"] + " | " + lv1 + " | " + lv2
    opex = ~df["final_capex_opex"].eq("Capex")

    for ent in ENTS:
        e = df[df.entity == ent]
        L.append(f"\n{'='*74}\n# {ent}")
        # §A opex H_LABOR 調整前/調整/後
        st = e[(e.horizontal_id == "H_LABOR") & opex.loc[e.index]]
        L.append(f"  §A opex 人工成本 by bucket（前/調整/後）：")
        g = st.groupby("yb")[["pre", "adj", "post"]].sum()
        for yb, r in g.iterrows():
            L.append(f"     {yb:<9} 前={r.pre:>9,.0f}  調整={r.adj:>9,.0f}  後={r.post:>9,.0f}")
        # §B 調整≠0 by (lv1,lv2) + H
        adj = e[e.adj.abs() > 0.5]
        if len(adj):
            L.append(f"  §B 有調整嘅行 by (調整一級|二級) → Σ調整 + 現H（top）：")
            ga = adj.assign(l1=lv1.loc[adj.index], l2=lv2.loc[adj.index]).groupby(["l1", "l2", "horizontal_label"])["adj"].sum().reset_index().sort_values("adj", key=lambda s: s.abs(), ascending=False)
            for _, r in ga.head(12).iterrows():
                L.append(f"     {str(r['l1'])[:14]:<14}|{str(r['l2'])[:14]:<14} H={r['horizontal_label'][:10]:<10} 調整={r['adj']:>9,.0f}")
        # §C staff 調整關鍵字
        kw = e[e["blob"].str.contains("支持部門|內部人工|內部資源|超過可計入|超出可計入|staff cost|payroll|人工成本", case=False, na=False, regex=True)]
        if len(kw):
            L.append(f"  §C staff/內部資源 關鍵字行 → 現H × bucket（Σ調整後，top）：")
            gk = kw.groupby(["horizontal_label", "yb"])["post"].sum().reset_index().sort_values("post", key=lambda s: s.abs(), ascending=False)
            for _, r in gk.head(10).iterrows():
                L.append(f"     {r['horizontal_label'][:12]:<12} {r['yb']:<9} 後={r['post']:>9,.0f}")
    _w()


def _w():
    out = ROOT / "results" / "inspect_staff_adj.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
