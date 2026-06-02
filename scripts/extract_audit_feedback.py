"""Extract + DEDUP the project-team audit feedback into compact rule views.

The audit workbook's "4_大表" sheet carries the project team's GROUND-TRUTH labels:
  U 原表類別       = project-team vertical (their category)
  V 原表科目分類   = project-team horizontal (their account class)
  W 大表類別匹配   = our vertical    ; X 大表科目匹配 = our horizontal
  Y 類別一致 / Z 科目一致 = agreement flags
  AA 類別分類備註 / AB 科目分類備註 = annotations (how to fix)
  AC = whether identifiable by ACC name
plus M project, N subproject, O account_code, P account_desc, C amount_mop.

27k annotated rows = only ~30 distinct annotations × repeats. This collapses them
to the DISTINCT rule tuples (the actual decisions), so the rulebase is small:

Outputs (results/):
  {stem}__mapping_sheet.tsv     — workbook 'mapping' sheet (official mapping, if any)
  {stem}__U_categories.tsv      — distinct 原表類別 (their V) × our V × amount  (master V map)
  {stem}__V_categories.tsv      — distinct 原表科目分類 (their H) × our H × amount (master H map)
  {stem}__vertical_rules.tsv    — (AA備註, 原表類別, our V) → n, amount, sample subproject
  {stem}__horizontal_rules.tsv  — (AB備註, account_code, account_desc, our H) → n, amount

Run (Windows):
  python scripts/extract_audit_feedback.py --file vml_audit_25.xlsx
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

# 4_大表 header names (stable)
C = dict(amt="amount_mop", their_V="原表類別", their_H="原表科目分類",
         our_V="大表類別匹配", our_H="大表科目匹配",
         vagree="類別一致", hagree="科目一致",
         aa="類別分類備註", ab="科目分類備註", acc_ok="是否可以直接根據ACC name識別出",
         proj="project", sub="subproject", code="account_code", adesc="account_desc")


def _amt(df, by):
    g = df.groupby(by, dropna=False, observed=True)[C["amt"]].agg(
        n="size", amount=lambda s: pd.to_numeric(s, errors="coerce").fillna(0).sum()).reset_index()
    return g.sort_values("amount", key=lambda s: s.abs(), ascending=False)


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--file", required=True)
    p.add_argument("--sheet", default="4_大表")
    args = p.parse_args()
    fp = Path(args.file)
    if not fp.exists():
        for alt in (Path("results")/args.file, Path.cwd().parent/args.file):
            if alt.exists(): fp = alt; break
    if not fp.exists():
        print(f"❌ {args.file} not found"); sys.exit(1)

    out = Path("results"); out.mkdir(exist_ok=True)
    stem = fp.stem
    xl = pd.ExcelFile(fp)
    print(f"file: {fp}\nsheets: {xl.sheet_names}")

    # mapping sheet (official)
    if "mapping" in xl.sheet_names:
        m = xl.parse("mapping")
        m.to_csv(out/f"{stem}__mapping_sheet.tsv", sep="\t", index=False, encoding="utf-8-sig")
        print(f"\n=== 'mapping' sheet ({len(m)} rows) → results/{stem}__mapping_sheet.tsv ===")
        print(m.head(40).to_csv(sep="\t", index=False))

    df = xl.parse(args.sheet)
    df.columns = [str(c).strip() for c in df.columns]
    for k in ("amt","their_V","their_H","our_V","our_H","aa","ab","code","adesc","sub"):
        if C[k] not in df.columns:
            print(f"  ⚠ column '{C[k]}' not found in {args.sheet}")
    print(f"\n[{args.sheet}] {len(df):,} rows")

    def S(col): return df[C[col]].astype(str).str.strip() if C[col] in df.columns else ""

    # 1) master category maps (their label × our label)
    if C["their_V"] in df.columns:
        uv = _amt(df.assign(**{C["their_V"]:S("their_V"), C["our_V"]:S("our_V")}),
                  [C["their_V"], C["our_V"]])
        uv.to_csv(out/f"{stem}__U_categories.tsv", sep="\t", index=False, encoding="utf-8-sig")
        print(f"\n=== 原表類別(their V) × our V — top by amount → {stem}__U_categories.tsv ===")
        print(uv.head(35).to_csv(sep="\t", index=False))
    if C["their_H"] in df.columns:
        vh = _amt(df.assign(**{C["their_H"]:S("their_H"), C["our_H"]:S("our_H")}),
                  [C["their_H"], C["our_H"]])
        vh.to_csv(out/f"{stem}__V_categories.tsv", sep="\t", index=False, encoding="utf-8-sig")
        print(f"\n=== 原表科目分類(their H) × our H — top by amount → {stem}__V_categories.tsv ===")
        print(vh.head(35).to_csv(sep="\t", index=False))

    # 2) vertical rules: (AA, their V, our V) + sample subproject
    if C["aa"] in df.columns:
        d = df[S("aa") != ""].copy()
        d2 = d.assign(**{C["aa"]:S("aa"), C["their_V"]:S("their_V"), C["our_V"]:S("our_V")})
        vr = d2.groupby([C["aa"], C["their_V"], C["our_V"]], dropna=False, observed=True).agg(
            n=(C["amt"],"size"),
            amount=(C["amt"], lambda s: pd.to_numeric(s,errors="coerce").fillna(0).sum()),
            sample_sub=(C["sub"], lambda s: " | ".join(pd.Series(s.astype(str)).dropna().unique()[:3])) if C["sub"] in df.columns else (C["amt"],"size"),
        ).reset_index().sort_values("amount", key=lambda s:s.abs(), ascending=False)
        vr.to_csv(out/f"{stem}__vertical_rules.tsv", sep="\t", index=False, encoding="utf-8-sig")
        print(f"\n=== VERTICAL rule view (AA × their V × our V) → {stem}__vertical_rules.tsv  ({len(vr)} rules) ===")
        print(vr.head(30).to_csv(sep="\t", index=False))

    # 3) horizontal rules: (AB, account_code, account_desc, our H)
    if C["ab"] in df.columns:
        d = df[S("ab") != ""].copy()
        d2 = d.assign(**{C["ab"]:S("ab"), C["code"]:S("code"), C["adesc"]:S("adesc"), C["our_H"]:S("our_H")})
        hr = d2.groupby([C["ab"], C["code"], C["adesc"], C["our_H"]], dropna=False, observed=True).agg(
            n=(C["amt"],"size"),
            amount=(C["amt"], lambda s: pd.to_numeric(s,errors="coerce").fillna(0).sum()),
        ).reset_index().sort_values("amount", key=lambda s:s.abs(), ascending=False)
        hr.to_csv(out/f"{stem}__horizontal_rules.tsv", sep="\t", index=False, encoding="utf-8-sig")
        print(f"\n=== HORIZONTAL rule view (AB × account_code × account_desc × our H) → {stem}__horizontal_rules.tsv  ({len(hr)} rules) ===")
        print(hr.head(30).to_csv(sep="\t", index=False))

    print("\n✓ dedup done — compact rule views in results/ (27k rows → distinct tuples)")


if __name__ == "__main__":
    main()
