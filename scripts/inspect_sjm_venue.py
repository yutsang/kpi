r"""sjm 24 venue/food split hint —— 食飲(ct=FnB 推廣費)嘅 description/subproject/成本中心/NG 有冇 venue 信號。
找「場地/活動/宴會/會展/Venue/Event/Hall」等可分 venue 嘅 row-level hint。

Run（Windows）:  python scripts\inspect_sjm_venue.py
出 results\inspect_sjm_venue.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
VEN_KW = r"場地|會場|宴會|活動|會展|展覽|演唱|表演|場館|租|Venue|Hall|Event|Banquet|Function|Ball|Concert|Show|Exhibit|Convention"


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    if not CSV.exists():
        print("!! csv 揾唔到"); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    e = df[df["entity"].astype(str) == "sjm"].copy()
    post = pd.to_numeric(e.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    e["m"] = post
    e["yb"] = e["year_bucket"].astype(str)
    for c in ("horizontal_id", "comp_type", "account_desc", "description", "subproject", "成本中心", "ng_code", "vertical_label"):
        e[c] = _s(e[c]) if c in e.columns else ""
    fnb = e[(e.yb == "24") & e.horizontal_id.eq("H_FNB")]
    L = [f"# sjm 24 食飲(H_FNB) Σ={fnb['m'].sum():,.0f}萬 — 揾 venue 信號（description/subproject/NG/V）"]
    L.append(f"\n── 食飲 by NG（golden NG3 會場2,266）──")
    for ng, g in fnb.groupby("ng_code"):
        L.append(f"   {ng:<5} Σ={g['m'].sum():,.0f}萬")
    L.append(f"\n── 食飲 by vertical_label（V 主題可能露 venue）──")
    gv = fnb.groupby("vertical_label")["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
    for _, r in gv.head(15).iterrows():
        L.append(f"   {r['vertical_label'][:24]:<24} {r['m']:>8,.0f}萬")
    L.append(f"\n── 食飲 by description top（有冇場地/活動字眼）──")
    gd = fnb.groupby("description")["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
    for _, r in gd.head(25).iterrows():
        hit = " ★venue?" if pd.Series([r["description"]]).str.contains(VEN_KW, regex=True, na=False).iloc[0] else ""
        L.append(f"   {r['description'][:40]:<40} {r['m']:>8,.0f}萬{hit}")
    L.append(f"\n── 食飲 by subproject top ──")
    gs = fnb.groupby("subproject")["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
    for _, r in gs.head(20).iterrows():
        hit = " ★venue?" if pd.Series([r["subproject"]]).str.contains(VEN_KW, regex=True, na=False).iloc[0] else ""
        L.append(f"   {r['subproject'][:40]:<40} {r['m']:>8,.0f}萬{hit}")
    ven = fnb[fnb["description"].str.contains(VEN_KW, regex=True, na=False) | fnb["subproject"].str.contains(VEN_KW, regex=True, na=False) | fnb["vertical_label"].str.contains(VEN_KW, regex=True, na=False)]
    L.append(f"\n── venre 信號命中合計 Σ={ven['m'].sum():,.0f}萬（golden 會場 3,938）──")
    p = ROOT / "results" / "inspect_sjm_venue.txt"; p.parent.mkdir(exist_ok=True)
    p.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"wrote {p.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
