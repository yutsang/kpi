"""dump_mgm_for_classify.py — prep MGM (company_6) for a Claude-bypass (no LLM):

  PART V  : the 140 unique projects (project sheet) — Project_code | Project_name | NG(Section) | scope(Category)
            → Claude assigns a vertical (V_*) per project → row_vertical_overrides column_map(Project_code→V).

  PART H  : the UNCOVERED horizontal sigs — i.e. combine rows whose H is NOT already decided by
            a WD/Source rule, an account-code predominant_rule, or a row_horizontal_override.
            These are the rows that would otherwise land empty / LLM. Grouped by (account_code, account_desc)
            with row-count + Σamt so Claude classifies the few that matter → predominant_rules.

Rows that DON'T need a sig (already decided) are reported as a coverage summary, not dumped.

  python scripts/dump_mgm_for_classify.py                 # both parts to stdout + results/*.tsv
  python scripts/dump_mgm_for_classify.py --top 60        # show top-60 uncovered H sigs inline
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parent.parent
CONF = ROOT / "conf" / "company_6" / "parameters.yml"
RAW = ROOT / "data" / "mgm" / "raw" / "mgm_25_raw.xlsx"
OUT = ROOT / "results"


def numify(s):
    return pd.to_numeric(s.astype(str).str.replace(",", "", regex=False).str.replace(r"^\s*-\s*$", "0", regex=True),
                         errors="coerce").fillna(0)


def _as_list(x):
    return [x] if isinstance(x, str) else (x or [])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", default=str(RAW))
    ap.add_argument("--top", type=int, default=50)
    a = ap.parse_args()
    if not Path(a.file).exists():
        print(f"X not found: {a.file}"); return
    OUT.mkdir(exist_ok=True)
    cf = yaml.safe_load(CONF.open(encoding="utf-8"))
    cols = cf["columns"]
    ACCT, DESC = cols["account_code"], cols["account_desc"]          # Ledger Account / Ledger Hierarchy Level 4
    AMT, SRC = cols["amount"], cols["capex_opex"]                    # Debit minus Credit / Source

    # ----- coverage rules from conf -----
    PRE = (cf.get("predominant_rules") or {}).get("horizontal") or []
    OVR = cf.get("row_horizontal_overrides") or []
    ovr_pref = [o["when"]["account_code_prefix"] for o in OVR
                if isinstance(o.get("when"), dict) and "account_code_prefix" in o["when"]]
    src_rule = {o["when"]["column_equals"]["value"]: o["set"] for o in OVR
                if isinstance(o.get("when"), dict) and isinstance(o["when"].get("column_equals"), dict)
                and o["when"]["column_equals"]["col"] == SRC}
    has_hlabel_map = any("column_map" in o and o["column_map"].get("col") == "H_Label" for o in OVR)

    def covered(acct, desc, source, hlabel):
        s = str(source).strip()
        if s in src_rule:
            return f"Source={s}"
        if has_hlabel_map and str(hlabel).strip() not in ("", "nan", "None"):
            return "H_Label"
        a, d = str(acct).strip(), str(desc).lower()
        for p in ovr_pref:
            if a.startswith(str(p)):
                return "override"
        for r in PRE:
            cond = r.get("if") or {}
            if any(a.startswith(str(p)) for p in _as_list(cond.get("account_code_prefix"))):
                return "rule"
            if a in [str(x) for x in _as_list(cond.get("account_code_equals"))]:
                return "rule"
            if any(str(k).lower() in d for k in _as_list(cond.get("account_desc_contains"))):
                return "rule"
        return ""

    xl = pd.ExcelFile(a.file)

    # ===== PART V : 140 projects =====
    proj = xl.parse("project")
    proj.columns = [str(c).strip() for c in proj.columns]
    pv = proj[["Project_code", "Project_name", "Category", "Section"]].copy()
    pv.to_csv(OUT / "mgm_v_projects.tsv", sep="\t", index=False)
    print(f"=== PART V — {len(pv)} projects (assign a V_* each) → results/mgm_v_projects.tsv ===")
    print("  Project_code\tNG(Section)\tscope\tProject_name")
    for _, r in pv.iterrows():
        print(f"  {str(r['Project_code'])[:14]:16s}{str(r['Section'])[:10]:12s}{str(r['Category'])[:11]:12s}{str(r['Project_name'])[:46]}")

    # ===== PART H : uncovered sigs =====
    df = xl.parse("combine")
    df.columns = [str(c).strip() for c in df.columns]
    for c in (ACCT, DESC, SRC):
        if c not in df.columns:
            df[c] = ""
    df["_amt"] = numify(df[AMT]) if AMT in df.columns else 0.0
    hl = "H_Label" if "H_Label" in df.columns else None
    df["_cov"] = [covered(ac, de, so, (df[hl].iloc[i] if hl else "")) for i, (ac, de, so) in
                  enumerate(zip(df[ACCT], df[DESC], df[SRC]))]
    cov = df.assign(_n=1).groupby(df["_cov"].replace("", "(UNCOVERED → needs sig)"))[["_n", "_amt"]].sum()
    print(f"\n=== coverage of {len(df):,} combine rows (H decided WITHOUT a sig vs not) ===")
    for k, r in cov.sort_values("_amt", key=abs, ascending=False).iterrows():
        print(f"  {str(k):28s} rows={int(r['_n']):7,}  Σ={r['_amt']:>16,.0f}")

    un = df[df["_cov"].eq("")]
    g = un.groupby([df[ACCT].astype(str).str.strip(), df[DESC].astype(str).str.strip()]).agg(
        n=("_amt", "size"), amt=("_amt", "sum")).reset_index()
    g.columns = ["account_code", "account_desc", "n", "amt"]
    g = g.reindex(g["amt"].abs().sort_values(ascending=False).index)
    g.to_csv(OUT / "mgm_h_uncovered.tsv", sep="\t", index=False)
    print(f"\n=== PART H — {len(g)} UNCOVERED (account_code, account_desc) sigs "
          f"(Σ={un['_amt'].sum():,.0f}, {len(un):,} rows) → results/mgm_h_uncovered.tsv ===")
    print("  account_code\taccount_desc\tn\tΣamt   (assign H_* each)")
    for _, r in g.head(a.top).iterrows():
        print(f"  {str(r['account_code'])[:14]:16s}{str(r['account_desc'])[:42]:44s}{int(r['n']):>6,}{r['amt']:>15,.0f}")
    if len(g) > a.top:
        print(f"  … +{len(g)-a.top} more in the TSV")


if __name__ == "__main__":
    main()
