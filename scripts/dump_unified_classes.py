"""Profile the NEW unified raw (same 27 cols for every entity) and dump the SMALL things
needed to (a) sanity-check the manual build and (b) plan the conf/tagging — WITHOUT moving
the data. Outputs are tiny (text + distinct values only) → never hits the 2MB email cap.

DEFAULT: loops EVERY data file in data/raw (xlsx/csv/tsv, skips ~$ temp files) and writes
TWO combined outputs:
  results/unified_ALL_profile.txt   ← per file: sheet names, cols missing/extra, blank%,
                                       year/entity/take_flag2/netoff distributions,
                                       unique_id dup check, amount tie
  results/unified_ALL_classes.tsv   ← per file: distinct pt_class_H / pt_class_V / ng_theme
                                       + rows + Σ adjusted (reference-mapping + 對數)
Paste BOTH back.

Run (Windows):
  python scripts/dump_unified_classes.py                      # loop data\raw
  python scripts/dump_unified_classes.py --file data\raw\vml_u_2025.xlsx   # just one
  python scripts/dump_unified_classes.py --entity vml         # loop, filter entity col
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
RAW_DIR = ROOT / "data" / "raw"
EXPECTED = ["unique_id", "entity", "year", "capex_opex", "account_code", "account_desc",
            "description", "project_code", "dicj_code", "project", "subproject", "ng_theme",
            "amount_mop", "adjustment_amount", "adjusted_amount", "adjust_lv1", "adjust_lv2",
            "pt_class_H", "pt_class_V", "vendor", "source", "comp_type", "is_labor",
            "is_internal", "take_flag2", "netoff_flag", "remark"]
KEY_CATS = ["year", "entity", "capex_opex", "ng_theme", "source", "comp_type",
            "is_labor", "is_internal", "take_flag2", "take_flag", "netoff_flag"]
SUFFIXES = (".xlsx", ".xlsm", ".xls", ".csv", ".tsv")


def _amt(df, col):
    if col not in df.columns:
        return pd.Series(0.0, index=df.index)
    return pd.to_numeric(df[col].astype("string").str.replace(",", "", regex=False),
                         errors="coerce").fillna(0.0)


def profile_one(path: Path, sheet, entity, top):
    """Returns (lines, classes_rows)."""
    L = [f"\n{'='*76}\n## FILE: {path.name}"]
    sheets = None
    if path.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        try:
            sheets = pd.ExcelFile(path).sheet_names
            L.append(f"   sheets: {sheets}  (reading {sheet or sheets[0]!r})")
        except Exception as e:
            L.append(f"   !! cannot open: {e}"); return L, []
        df = pd.read_excel(path, sheet_name=sheet or 0, dtype=str)
    else:
        df = pd.read_csv(path, sep="\t" if path.suffix.lower() == ".tsv" else ",", dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    if entity and "entity" in df.columns:
        df = df[df["entity"].astype("string").str.strip().str.lower() == entity.lower()]
    L.append(f"   rows={len(df):,}")

    miss = [c for c in EXPECTED if c not in df.columns]
    extra = [c for c in df.columns if c not in EXPECTED]
    L.append(f"   MISSING cols ({len(miss)}): {miss}")
    L.append(f"   EXTRA   cols ({len(extra)}): {extra}")

    adj = _amt(df, "adjusted_amount")
    am, ad = _amt(df, "amount_mop"), _amt(df, "adjustment_amount")
    if "unique_id" in df.columns:
        s = df["unique_id"].astype("string").fillna("").str.strip()
        L.append(f"   unique_id: non-blank={s.ne('').sum():,}  distinct={s[s.ne('')].nunique():,}"
                 f"  → {'ROW-UNIQUE' if s[s.ne('')].nunique() >= s.ne('').sum() else 'document-level (repeats)'}")
    L.append(f"   Σamount_mop={am.sum():,.0f}  Σadjustment={ad.sum():,.0f}  Σadjusted={adj.sum():,.0f}"
             f"  | rows adjusted≠mop+adj: {int((((am+ad)-adj).abs() > 1).sum()):,}")

    for c in KEY_CATS:
        if c not in df.columns:
            L.append(f"   -- {c}: MISSING"); continue
        s = df[c].astype("string").fillna("").str.strip()
        vc = pd.DataFrame({"k": s.replace("", "(blank)"), "amt": adj}).groupby("k").agg(
            n=("amt", "size"), amt=("amt", "sum")).sort_values("amt", key=lambda x: x.abs(), ascending=False)
        tops = "  ".join(f"{k}={int(r['n'])}r/{r['amt']/1e6:.1f}M" for k, r in vc.head(8).iterrows())
        L.append(f"   -- {c}: blank={s.eq('').mean()*100:.0f}%  distinct={s.nunique()}  | {tops}")

    rows = []
    for axis in ("pt_class_V", "pt_class_H", "ng_theme"):
        if axis not in df.columns: continue
        s = df[axis].astype("string").fillna("").str.strip().replace("", "(blank)")
        g = pd.DataFrame({"v": s, "amt": adj}).groupby("v").agg(n=("amt", "size"), amt=("amt", "sum"))
        for v, r in g.sort_values("amt", key=lambda x: x.abs(), ascending=False).iterrows():
            rows.append({"file": path.name, "axis": axis, "value": v,
                         "rows": int(r["n"]), "amount": round(r["amt"])})
    return L, rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dir", default=str(RAW_DIR))
    ap.add_argument("--file", default=None, help="profile just this one file")
    ap.add_argument("--sheet", default=None)
    ap.add_argument("--entity", default=None)
    ap.add_argument("--top", type=int, default=40)
    a = ap.parse_args()

    if a.file:
        files = [Path(a.file)]
    else:
        d = Path(a.dir)
        files = sorted(p for p in d.iterdir()
                       if p.suffix.lower() in SUFFIXES and not p.name.startswith("~$"))
        if not files:
            print(f"X no data files found in {d}"); return
    L = [f"# dump_unified_classes — {len(files)} file(s) from {a.dir}"]
    all_rows = []
    for p in files:
        try:
            lines, rows = profile_one(p, a.sheet, a.entity, a.top)
        except Exception as e:
            lines, rows = [f"\n{'='*76}\n## FILE: {p.name}\n   !! ERROR: {e}"], []
        L += lines; all_rows += rows

    outdir = ROOT / "results"; outdir.mkdir(exist_ok=True)
    txt = outdir / "unified_ALL_profile.txt"
    tsv = outdir / "unified_ALL_classes.tsv"
    txt.write_text("\n".join(L), encoding="utf-8")
    pd.DataFrame(all_rows).to_csv(tsv, sep="\t", index=False)
    print("\n".join(L))
    print(f"\nwrote {txt.relative_to(ROOT)}  ({txt.stat().st_size/1e3:.0f} KB)")
    print(f"wrote {tsv.relative_to(ROOT)}  ({tsv.stat().st_size/1e3:.0f} KB)   ← paste BOTH back")


if __name__ == "__main__":
    main()
