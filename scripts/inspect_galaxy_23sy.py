"""Inspect galaxy_2024_23sy.xlsx (the 24_23SY supplement) to find WHY H_OTHER=58.6% + V空=28.3%.
The dominant source (9.ADR ~43,801 rows) is calculated data with NO SAP account_code/desc, so the
SAP-prefix H rules + the 基礎|一級/二級標簽 column_map (which belong to the 2025 main file) can't
fire. This dump finds the supplement's OWN label column so we can wire a column_map / Source fallback.

Run (Windows):  python scripts/inspect_galaxy_23sy.py [--pw dicj_kpmg]
Output: prints + results/inspect_galaxy_23sy.txt
"""
from __future__ import annotations
import argparse, io, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
LABEL_HINT = ("標簽", "標籤", "分類", "類別", "Nature", "Source", "System", "範疇", "一級", "二級")


def find_file():
    for p in [ROOT/"data"/"galaxy"/"raw"/"galaxy_2024_23sy.xlsx",
              ROOT/"data"/"galaxy"/"raw"/"2023"/"galaxy_2024_23sy.xlsx"]:
        if p.exists(): return p
    hits = list((ROOT/"data").rglob("galaxy_2024_23sy.xlsx"))
    return hits[0] if hits else None


def load(fp, pw):
    if pw:
        try:
            import msoffcrypto
            buf = io.BytesIO()
            with open(fp, "rb") as f:
                of = msoffcrypto.OfficeFile(f); of.load_key(password=pw); of.decrypt(buf)
            buf.seek(0)
            return pd.read_excel(buf, sheet_name="Combine", header=1, dtype=object)
        except Exception as e:
            print(f"  decrypt failed ({e}) — trying plain");
    return pd.read_excel(fp, sheet_name="Combine", header=1, dtype=object)


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--pw", default=None); a = ap.parse_args()
    fp = find_file()
    if not fp: print("X galaxy_2024_23sy.xlsx not found under data/galaxy/raw/"); return
    print(f"file: {fp}")
    df = load(fp, a.pw)
    L = [f"# inspect_galaxy_23sy  ({len(df):,} rows)", f"\n## ALL columns ({len(df.columns)}):", f"   {list(df.columns)}"]

    def find(*subs):
        for c in df.columns:
            cl = str(c).lower()
            if any(s.lower() in cl for s in subs): return c
        return None

    amtc = find("amount") or "Amount"
    srcc = find("source", "system", "來源") or None
    a_amt = pd.to_numeric(df[amtc], errors="coerce").fillna(0.0) if amtc in df.columns else pd.Series(0.0, index=df.index)
    atot = a_amt.abs().sum() or 1
    L.append(f"\n## amount={amtc!r}  Σ={a_amt.sum():,.0f}  source_col={srcc!r}")

    if srcc:
        L.append(f"\n## '{srcc}' distribution (rows | Σ|amt|%):")
        for v, g in df.groupby(df[srcc].astype(str)):
            L.append(f"   {str(v)[:34]:34s} {len(g):>7,}  {a_amt.abs()[g.index].sum()/atot*100:5.1f}%")

    # every column whose name hints a label/category → value_counts (this is what column_map needs)
    L.append("\n## candidate LABEL columns (name contains 標簽/分類/類別/Nature/範疇/一級/二級):")
    for c in df.columns:
        if any(h in str(c) for h in LABEL_HINT):
            s = df[c].astype("string").fillna("").str.strip()
            nun = s[s.ne("")].nunique()
            L.append(f"\n   [{c}]  distinct={nun}  blank%={s.eq('').mean()*100:.0f}")
            for v, amt_s in a_amt.abs().groupby(s).sum().sort_values(ascending=False).head(20).items():
                if v == "": continue
                L.append(f"      {str(v)[:38]:38s} {amt_s/atot*100:5.1f}%")

    # ADR-subset (or biggest source): are account_code/desc actually blank?
    accc = find("account code", "account_code", "科目代")
    adc = find("account desc", "account_desc", "科目描", "摘要")
    projc = find("project", "投資項目", "項目名")
    ngc = find("ng11", "ng category", "ng 分")
    if srcc:
        big = df[srcc].astype(str).value_counts().index[0]
        sub = df[df[srcc].astype(str).eq(big)]
        L.append(f"\n## biggest source = {big!r} ({len(sub):,} rows) — col fill rates:")
        for nm, c in [("account_code", accc), ("account_desc", adc), ("project", projc), ("ng11", ngc)]:
            if c and c in sub.columns:
                ss = sub[c].astype("string").fillna("").str.strip()
                top = " | ".join(map(str, ss[ss.ne('')].value_counts().head(6).index))
                L.append(f"   {nm:12s} ({c}): nonblank={ss.ne('').mean()*100:.0f}%  top: {top[:90]}")
            else:
                L.append(f"   {nm:12s}: column not found")

    out = ROOT/"results"/"inspect_galaxy_23sy.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
