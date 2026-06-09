"""Prebuild SJM 2023 raw — stack the two differently-shaped JE tabs (項目明細賬capex +
項目明細賬opex) into ONE sheet with SJM's pipeline column names, so yearly_sources 2023 can
point at it. Gross (調整 reconciled later). unique_id = hash (no project name).

capex tab: Val/COArea Crcy | Project Code with Project Name | Cost Element | Cost element descr. | 項目性質
opex  tab: Amount in local currency | 项目名称 | G/L Account | GL account description | 項目性質
→ common: Val/COArea Crcy | Project Name | Cost Element | Cost element descr. | 項目性質 | Capex/Opex | Fiscal Year

Run (Windows):
  python scripts/build_sjm_23.py
Output → data/sjm/raw/2023/sjm_23_raw.xlsx  (sheet 'combine')
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent

TABS = [
    dict(sheet="項目明細賬capex", capex="Capex", amt="Val/COArea Crcy",
         proj="Project Code with Project Name", ac="Cost Element", ad="Cost element descr.", v="項目性質",
         desc="Description", vendor="AC1 Vendor Name", wbs="CO object name"),
    dict(sheet="項目明細賬opex", capex="Opex", amt="Amount in local currency",
         proj="项目名称", ac="G/L Account", ad="GL account description", v="項目性質",
         desc="Description", vendor="AC1 vendor name", wbs="WBS Name"),
]


def find_file(name):
    p = Path(name)
    if p.exists(): return p
    d = ROOT / "data"
    hits = list(d.glob(f"*/raw/{name}")) + list(d.rglob(name)) + list(d.rglob(f"*{name}*"))
    exact = [h for h in hits if h.name == name]
    return (exact or hits)[0] if (exact or hits) else None


def col(df, name):
    if name in df.columns: return name
    for c in df.columns:
        if str(c).strip() == str(name).strip(): return c
    return None


def num(s): return pd.to_numeric(s, errors="coerce")


def main():
    fp = find_file("sjm_2023.xlsx")
    if not fp:
        print("X sjm_2023.xlsx not found under data/*/raw/"); return
    out = []
    for t in TABS:
        try:
            df = pd.read_excel(fp, sheet_name=t["sheet"], header=0, dtype=object)
        except Exception as e:
            print(f"[{t['sheet']}] read failed: {e}"); continue
        amt, proj, ac, ad, v = (col(df, t[k]) for k in ("amt", "proj", "ac", "ad", "v"))
        de, ve, wb = (col(df, t.get(k, "")) for k in ("desc", "vendor", "wbs"))
        if not amt: print(f"[{t['sheet']}] amount {t['amt']!r} missing"); continue
        _clean = lambda c: df[c].astype(str).str.replace(r"[\r\n]+", " ", regex=True).str.strip().values
        sub = pd.DataFrame({
            "Val/COArea Crcy": num(df[amt]).values,
            "Project Name": _clean(proj) if proj else "",
            "Cost Element": df[ac].astype(str).str.strip().values if ac else "",
            "Cost element descr.": df[ad].astype(str).str.strip().values if ad else "",
            "Description": _clean(de) if de else "",            # 摘要（推廣費可按此 row 拆）
            "AC1 Vendor Name": df[ve].astype(str).str.strip().values if ve else "",  # 供應商
            "WBS Name": _clean(wb) if wb else "",               # 用途明細（opex WBS Name / capex CO object name）
            "項目性質": df[v].astype(str).str.strip().values if v else "",
            "Capex/Opex": t["capex"],
            "Fiscal Year": 2023,
        })
        out.append(sub)
        print(f"[{t['sheet']}] {len(sub):,} rows  Σ={sub['Val/COArea Crcy'].sum():,.0f}")
    if not out: print("nothing built"); return
    allr = pd.concat(out, ignore_index=True)
    allr["unique_id"] = ["S23_" + str(i) for i in range(len(allr))]
    res = ROOT / "data" / "sjm" / "raw"; res.mkdir(parents=True, exist_ok=True)   # step0 reads from raw/
    fpo = res / "sjm_23_raw.xlsx"
    with pd.ExcelWriter(fpo, engine="openpyxl") as w:
        allr.to_excel(w, sheet_name="combine", index=False)
    print(f"\n✓ {len(allr):,} rows → {fpo.relative_to(ROOT)}  Σ={allr['Val/COArea Crcy'].sum():,.0f}")
    print(allr.groupby("Capex/Opex")["Val/COArea Crcy"].sum().round(0).to_string())


if __name__ == "__main__":
    main()
