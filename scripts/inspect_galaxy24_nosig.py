"""Galaxy — content of the blank-acct rows (the 434M S4 cluster) so the source→H mapping
can be decided from EVIDENCE. User: 呢啲行成個 source 係同一分類；SAP0/1/2 有分類睇內容；
PMS=comp；10.HK Limo=其他(非Comp其他)。

For every galaxy file: rows where account_code AND account_desc are blank →
group by source × pt_class_H: rows, Σ調整後(無調整欄則調整前), sample descriptions.

Run (Windows):  python scripts/inspect_galaxy24_nosig.py
Output: prints + results/inspect_galaxy24_nosig.txt (small — paste back)
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "data" / "raw"
USE = ["account_code", "account_desc", "description", "source", "pt_class_H", "pt_class_V",
       "comp_type", "amount_mop", "adjustment_amount", "adjusted_amount", "year"]


def _num(s):
    return pd.to_numeric(s.astype("string").str.replace(",", "", regex=False), errors="coerce")


def _s(df, col):
    return df.get(col, pd.Series("", index=df.index)).astype("string").fillna("").str.strip()


def main():
    L = ["# inspect_galaxy24_nosig — blank-acct rows: source × pt_class_H content (S4 434M cluster)"]
    for f in sorted(RAW.glob("galaxy_*.xls*")):
        if f.name.startswith("~$"):
            continue
        sheets = pd.ExcelFile(f).sheet_names
        sh = next((s for s in sheets if str(s).strip().lower() == "data"), sheets[0])
        hdr = pd.read_excel(f, sheet_name=sh, nrows=0)
        use = [c for c in USE if c in [str(x).strip() for x in hdr.columns]]
        df = pd.read_excel(f, sheet_name=sh, dtype=str, usecols=use)
        df.columns = [str(c).strip() for c in df.columns]
        blank = _s(df, "account_code").eq("") & _s(df, "account_desc").eq("")
        sub = df[blank].copy()
        fin = _num(_s(sub, "adjusted_amount"))
        fin = fin.fillna(_num(_s(sub, "amount_mop")).fillna(0.0) + _num(_s(sub, "adjustment_amount")).fillna(0.0))
        L.append(f"\n{'='*72}\n## {f.name} (sheet={sh!r}) — blank-acct rows: {len(sub):,} / {len(df):,}"
                 f"  Σ={fin.sum()/1e6:,.2f}M")
        if not len(sub):
            continue
        src, pch = _s(sub, "source"), _s(sub, "pt_class_H")
        de = _s(sub, "description")
        g = (pd.DataFrame({"src": src.replace("", "(blank)"), "pch": pch.replace("", "(blank)"), "amt": fin})
             .groupby(["src", "pch"])["amt"].agg(["sum", "size"])
             .sort_values("sum", key=lambda x: x.abs(), ascending=False))
        L.append(f"   {'source':28s} {'pt_class_H':26s} {'rows':>6s} {'Σ(M)':>10s}")
        for (s_, p_), r in g.head(25).iterrows():
            L.append(f"   {str(s_)[:28]:28s} {str(p_)[:26]:26s} {int(r['size']):>6,} {r['sum']/1e6:>10.2f}")
        L.append("   -- sample descriptions per source:")
        for s_ in src.replace("", "(blank)").value_counts().head(8).index:
            m = src.replace("", "(blank)").eq(s_)
            samples = de[m][de[m].ne("")].drop_duplicates().head(3).tolist()
            L.append(f"      {str(s_)[:24]:24s} {[str(x)[:46] for x in samples]}")
    out = ROOT / "results" / "inspect_galaxy24_nosig.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste back")


if __name__ == "__main__":
    main()
