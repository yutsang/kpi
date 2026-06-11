"""Profile the NEW unified raw (one file, same 27 cols for every entity) and dump the SMALL
things needed to (a) sanity-check the manual build and (b) build the pt_class → our-taxonomy map
WITHOUT re-running the LLM. Output is tiny (distinct values only) → never hits the 2MB email cap.

What it prints / writes:
  • columns present vs expected (missing / extra)
  • per key-col: blank% + distinct count + top values (rows + Σadjusted_amount)
  • unique_id: rows vs distinct (flags if NOT row-unique — repeats = document-level)
  • amount tie: Σamount_mop vs Σadjusted_amount, and rows where adjusted ≠ amount_mop+adjustment
  • take_flag2 / netoff_flag value distribution (Σ amount each) — which rows are INCLUDED
  • results/unified_<ent>_classes.tsv : distinct pt_class_H / pt_class_V / ng_theme + count + Σamt
        (THIS is the file I use to write the column_map; paste it back — it is < ~100 KB)

Run (Windows):
  python scripts/dump_unified_classes.py --file data\unified\vml.xlsx
  python scripts/dump_unified_classes.py --file data\unified\all.csv --entity vml   # filter one entity from a combined file
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
EXPECTED = ["unique_id", "entity", "year", "capex_opex", "account_code", "account_desc",
            "description", "project_code", "dicj_code", "project", "subproject", "ng_theme",
            "amount_mop", "adjustment_amount", "adjusted_amount", "adjust_lv1", "adjust_lv2",
            "pt_class_H", "pt_class_V", "vendor", "source", "comp_type", "is_labor",
            "is_internal", "take_flag2", "netoff_flag", "remark"]
KEY_CATS = ["year", "entity", "capex_opex", "ng_theme", "pt_class_H", "pt_class_V", "source",
            "comp_type", "is_labor", "is_internal", "take_flag2", "netoff_flag"]


def _read(path, sheet):
    p = Path(path)
    if p.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        return pd.read_excel(p, sheet_name=sheet or 0, dtype=str)
    return pd.read_csv(p, sep="\t" if p.suffix.lower() == ".tsv" else ",", dtype=str)


def _amt(df, col):
    if col not in df.columns:
        return pd.Series(0.0, index=df.index)
    return pd.to_numeric(df[col].astype("string").str.replace(",", "", regex=False), errors="coerce").fillna(0.0)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", required=True)
    ap.add_argument("--sheet", default=None)
    ap.add_argument("--entity", default=None)
    ap.add_argument("--top", type=int, default=40)
    a = ap.parse_args()
    df = _read(a.file, a.sheet)
    df.columns = [str(c).strip() for c in df.columns]
    if a.entity and "entity" in df.columns:
        df = df[df["entity"].astype("string").str.strip().str.lower() == a.entity.lower()]
    ent = a.entity or (df["entity"].dropna().iloc[0] if "entity" in df.columns and len(df) else "unified")
    stem = Path(a.file).stem + (f"_{a.entity}" if a.entity else "")   # per-file output names (3 files/entity must not overwrite)
    L = [f"# dump_unified_classes — {ent}  rows={len(df):,}  file={a.file}"]

    miss = [c for c in EXPECTED if c not in df.columns]
    extra = [c for c in df.columns if c not in EXPECTED]
    L.append(f"\n## columns: {len(df.columns)} present")
    L.append(f"   MISSING ({len(miss)}): {miss}")
    L.append(f"   EXTRA   ({len(extra)}): {extra}")

    adj = _amt(df, "adjusted_amount")
    L.append("\n## unique_id traceability:")
    if "unique_id" in df.columns:
        nun = df["unique_id"].astype('string').fillna("").ne("").sum()
        dis = df["unique_id"].nunique(dropna=True)
        L.append(f"   rows={len(df):,}  non-blank={nun:,}  distinct={dis:,}  "
                 f"→ {'ROW-UNIQUE ✓' if dis >= len(df) else f'NOT row-unique (repeats; document-level)'}")
    L.append("\n## amount tie:")
    am, ad = _amt(df, "amount_mop"), _amt(df, "adjustment_amount")
    L.append(f"   Σamount_mop={am.sum():,.0f}  Σadjustment={ad.sum():,.0f}  Σadjusted_amount={adj.sum():,.0f}")
    bad = ((am + ad) - adj).abs() > 1
    L.append(f"   rows where adjusted ≠ amount_mop+adjustment: {int(bad.sum()):,}")

    for c in KEY_CATS:
        if c not in df.columns:
            L.append(f"\n## {c}: MISSING"); continue
        s = df[c].astype("string").fillna("").str.strip()
        blank = s.eq("").mean() * 100
        vc = pd.DataFrame({"k": s.replace("", "(blank)"), "amt": adj}).groupby("k").agg(
            n=("amt", "size"), amt=("amt", "sum")).sort_values("amt", key=lambda x: x.abs(), ascending=False)
        L.append(f"\n## {c}: blank={blank:.0f}%  distinct={s.nunique()} ")
        for k, r in vc.head(a.top).iterrows():
            L.append(f"   {str(k)[:34]:34s} {int(r['n']):>7} rows  {r['amt']/1e6:>10.3f}M")

    # small TSV for mapping (distinct pt_class_H / pt_class_V / ng_theme)
    out = ROOT / "results" / f"unified_{stem}_classes.tsv"; out.parent.mkdir(exist_ok=True)
    rows = []
    for axis in ("pt_class_V", "pt_class_H", "ng_theme"):
        if axis not in df.columns: continue
        s = df[axis].astype("string").fillna("").str.strip().replace("", "(blank)")
        g = pd.DataFrame({"v": s, "amt": adj}).groupby("v").agg(n=("amt", "size"), amt=("amt", "sum"))
        for v, r in g.sort_values("amt", key=lambda x: x.abs(), ascending=False).iterrows():
            rows.append({"axis": axis, "value": v, "rows": int(r["n"]), "amount_mop": round(r["amt"]), "our_tag": ""})
    pd.DataFrame(rows).to_csv(out, sep="\t", index=False)
    txt = ROOT / "results" / f"unified_{stem}_profile.txt"
    txt.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {txt.relative_to(ROOT)}  +  {out.relative_to(ROOT)} (paste this one back)")


if __name__ == "__main__":
    main()
