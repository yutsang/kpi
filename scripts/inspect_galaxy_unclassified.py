r"""inspect_galaxy_unclassified.py — galaxy「未分類項目」係咩、點解 DICJ 空
Run: python scripts\inspect_galaxy_unclassified.py
In : tableau_combined_25.csv
Out: results\inspect_galaxy_unclassified.txt
拆解 entity=galaxy 且 project=未分類項目 嘅行：by NG/V/H/capex-opex/year + top account_desc（行數+金額萬）。
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT  = ROOT / "results" / "inspect_galaxy_unclassified.txt"
CSV  = next((c for c in [ROOT/"tableau_combined_25.csv", Path("tableau_combined_25.csv")] if c.exists()), None)
COLS = ["entity", "project", "dicj code", "amount_mop", "year_bucket", "ng_label", "vertical_label",
        "horizontal_label", "final_capex_opex", "account_code", "account_desc"]


def main():
    if CSV is None:
        OUT.parent.mkdir(exist_ok=True); OUT.write_text("!! csv 揾唔到", encoding="utf-8"); print("no csv"); return
    acc = {}
    tot_rows = 0; tot_amt = 0.0
    for ch in pd.read_csv(CSV, usecols=lambda c: c in COLS, chunksize=300_000, dtype=str, encoding="utf-8-sig"):
        ch = ch.rename(columns=lambda c: c.strip())
        ch = ch[(ch["entity"] == "galaxy") & (ch["project"].astype(str).str.contains("未分類", na=False))]
        if not len(ch): continue
        ch["amt"] = pd.to_numeric(ch["amount_mop"], errors="coerce").fillna(0.0) / 1e4
        ch["rows"] = 1
        for c in COLS:
            if c in ch and c != "amount_mop": ch[c] = ch[c].fillna("").astype(str).str.strip()
        tot_rows += len(ch); tot_amt += ch["amt"].sum()
        for key, by in [("co","final_capex_opex"),("yr","year_bucket"),("ng","ng_label"),
                        ("v","vertical_label"),("h","horizontal_label"),
                        ("ad",["account_code","account_desc"])]:
            bys = by if isinstance(by, list) else [by]
            acc.setdefault(key, []).append(ch.groupby(bys, as_index=False).agg(amt=("amt","sum"), rows=("rows","sum")))
        print(f"  matched: {tot_rows:,} 行 / {tot_amt:,.0f}萬", flush=True)

    L = [f"# galaxy 未分類項目 拆解", f"總計：{tot_rows:,} 行 / {tot_amt:,.0f}萬", ""]
    def dump(key, title, topn=30):
        parts = acc.get(key, [])
        if not parts: return
        df = pd.concat(parts, ignore_index=True)
        keys = [c for c in df.columns if c not in ("amt", "rows")]
        g = df.groupby(keys, as_index=False).agg(amt=("amt","sum"), rows=("rows","sum")).sort_values("amt", ascending=False)
        L.append(f"── {title} ──")
        for _, r in g.head(topn).iterrows():
            k = " | ".join(str(r[c]) for c in keys)
            L.append(f"   {k:<48} {r['amt']:>11,.0f}萬 {int(r['rows']):>9,}行")
        L.append("")
    dump("co","by capex/opex"); dump("yr","by year_bucket"); dump("ng","by NG label")
    dump("v","by vertical (V)"); dump("h","by horizontal (H)"); dump("ad","top account_code+desc", 40)
    OUT.parent.mkdir(exist_ok=True); OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {OUT.relative_to(ROOT)}")

if __name__ == "__main__":
    main()
