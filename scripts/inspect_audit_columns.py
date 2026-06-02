"""Inspect the audit workbook's columns + our pipeline columns, to wire PBC rules.

The VML-25 rule spec depends on PBC helper columns (Payroll flag, comp-nature,
項目性質, NG-limit, subproject). Those live in the project-team audit Excel but may
NOT flow into our pipeline's tagged_rows. Before writing any rule we must know:
  1. the EXACT header text + spreadsheet letter of each helper column, per sheet
  2. whether the same column reaches data/{ent}/interim/{com}_tagged_rows.parquet
  3. if not, what join key (account_code + project + subproject) can bridge them

This dumps, for every sheet of the audit file:
  col-index → spreadsheet letter (A,B,..,AA) → header → n_unique → sample values
and flags candidate helper columns by keyword. Then (if --entity given) dumps our
parquet columns the same way, so the two can be lined up.

Outputs (results/):
  {stem}__columns.tsv          — every sheet × column (letter, header, n_uniq, samples)
  {ent}_pipeline_columns.tsv   — our tagged_rows columns (if --entity given)

Run (Windows):
  python scripts/inspect_audit_columns.py --file vml_audit_25.xlsx --entity vml
"""
from __future__ import annotations
import argparse, sys, csv
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
import pandas as pd

ENT = {"galaxy":"company_1","sjm":"company_2","wynn":"company_3",
       "vml":"company_4","melco":"company_5","mgm":"company_6"}

# headers hinting a helper/PBC column we rely on
HELPER_HINTS = {
    "payroll":  ["payroll", "人工", "工資", "薪"],
    "comp_nat": ["comp", "complimentary", "性質", "nature", "c&e", "arena", "hotel room"],
    "proj_nat": ["項目性質", "性質", "nature of"],
    "ng_limit": ["ng", "ng_limit", "limit", "範圍", "range"],
    "subproj":  ["subproject", "sub project", "sub_project", "子項", "細項"],
    "project":  ["project", "項目", "工程"],
}


def col_letter(i: int) -> str:
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def flag(header: str) -> str:
    h = str(header).lower()
    hits = [role for role, kws in HELPER_HINTS.items() if any(k in h for k in kws)]
    return "/".join(hits)


def dump_df(df: pd.DataFrame, label: str, writer, console_max=60):
    print(f"\n=== {label} — {len(df):,} rows × {df.shape[1]} cols ===")
    print(f"  {'letter':<6} {'header':<40} {'n_uniq':>7}  flag        sample")
    for i, c in enumerate(df.columns):
        try:
            vals = df[c].dropna().astype(str)
            uniq = vals[vals.str.strip() != ""].unique()[:3]
            n = df[c].nunique(dropna=True)
        except Exception:
            uniq, n = [], 0
        fl = flag(c)
        samp = " | ".join(str(x)[:22] for x in uniq)
        writer.writerow([label, col_letter(i), str(c), n, fl, samp])
        if i < console_max:
            mark = "  ⭐" if fl else "    "
            print(f"{mark}{col_letter(i):<6} {str(c)[:40]:<40} {n:>7}  {fl:<10}  {samp[:48]}")


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--file", required=True)
    p.add_argument("--entity", default=None, choices=list(ENT))
    args = p.parse_args()

    fp = Path(args.file)
    if not fp.exists():
        for alt in (Path("results")/args.file, Path.cwd().parent/args.file):
            if alt.exists(): fp = alt; break
    if not fp.exists():
        print(f"X {args.file} not found"); sys.exit(1)

    out = Path("results"); out.mkdir(exist_ok=True)
    stem = fp.stem
    xl = pd.ExcelFile(fp)
    print(f"file: {fp}\nsheets: {xl.sheet_names}")

    with (out/f"{stem}__columns.tsv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter="\t")
        w.writerow(["sheet", "letter", "header", "n_uniq", "helper_flag", "samples"])
        for sh in xl.sheet_names:
            try:
                df = xl.parse(sh, nrows=4000)
            except Exception as e:
                print(f"  [skip {sh}] {e}"); continue
            df.columns = [str(c).strip() for c in df.columns]
            dump_df(df, sh, w)
    print(f"\n-> results/{stem}__columns.tsv (all sheets)")

    if args.entity:
        com = ENT[args.entity]
        pq = Path(f"data/{args.entity}/interim/{com}_tagged_rows.parquet")
        if pq.exists():
            df = pd.read_parquet(pq)
            with (out/f"{args.entity}_pipeline_columns.tsv").open("w", encoding="utf-8-sig", newline="") as f:
                w = csv.writer(f, delimiter="\t")
                w.writerow(["sheet", "letter", "header", "n_uniq", "helper_flag", "samples"])
                dump_df(df.head(4000), f"PIPELINE:{args.entity}_tagged_rows", w)
            print(f"\n-> results/{args.entity}_pipeline_columns.tsv (our pipeline)")
            print("\n  >> compare the ⭐ helper columns above: which exist in BOTH the audit")
            print("     file AND our pipeline? Those are auto-applicable. The rest need a join")
            print("     (account_code + project + subproject) or a new PBC input.")
        else:
            print(f"\n  (our pipeline parquet {pq} not found — run kedro through step4 first)")


if __name__ == "__main__":
    main()
