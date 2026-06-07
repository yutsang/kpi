"""Prebuild MGM 2023 raw — union the ROW-LEVEL JE sheets into ONE sheet (2025 'combine'
schema) so the pipeline tags V/H per 項目. build_mgm_golden then distributes the Leadsheet
golden (866M) by these V/H proportions (gaming capex 1.337B is pivot-only → golden-driven).

2025 schema cols: Debit minus Credit | Source | Project_code | Section.1 | Ledger Account
                  | Ledger Hierarchy Level 5

Row-level sources (2023):
  OPEX.xlsx  WD#1/WD#2/WD#3 (Workday opex, keyed 承批公司項目序號) · PM (Patron comp) ·
             Data#4/Data#5 (payroll) · Data#6 (COGS)
  CAPEX.xlsx JL details (non-gaming capex JE, Project Code)
  MGM-gaming.xlsx 項目1..項目10 (gaming capex, tab = 項目序號)
NG (Section.1) = 項目序號 → 項目性質 from Leadsheet; gaming → 博彩娛樂.

Run (Windows):
  python scripts/build_mgm_23_raw.py
Output → data/mgm/raw/mgm_23_raw.xlsx  (sheet 'combine')
"""
from __future__ import annotations
import re, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
YDIR = ROOT / "data" / "mgm" / "raw" / "2023"


def find_file(name):
    for p in [YDIR / name, ROOT / "data" / "mgm" / "raw" / name]:
        if p.exists(): return p
    hits = list((ROOT / "data").rglob(name))
    return hits[0] if hits else None


def col(df, name):
    if name is None: return None
    if name in df.columns: return name
    for c in df.columns:
        if str(c).strip() == str(name).strip(): return c
    for c in df.columns:                       # fuzzy contains
        if str(name).strip() in str(c).strip():
            return c
    return None


def num(s): return pd.to_numeric(s, errors="coerce").fillna(0.0)


def seq(v):
    m = re.search(r"\d+", str(v))
    return "項目" + m.group(0) if m else ""


# (file, sheet, header, amount, project_col, account, hier, source, capex, ng_const)
SOURCES = [
    ("OPEX.xlsx",       "WD#1",       0, "Amount",             "承批公司項目序號", "Ledger Account", "Ledger Hierarchy Level 4", "WD1",     "Opex", None),
    ("OPEX.xlsx",       "WD#2",       0, "Amount",             "承批公司項目序號", "Ledger Account", "Ledger Hierarchy Level 4", "WD2",     "Opex", None),
    ("OPEX.xlsx",       "WD#3",       0, "Amount",             "承批公司項目序號", "Ledger Account", "Ledger Hierarchy Level 4", "WD3",     "Opex", None),
    ("OPEX.xlsx",       "PM",         0, "MOP",                "承批公司項目序號", "Item Type",      "Promotion Type",           "PM",      "Opex", None),
    ("OPEX.xlsx",       "Data#4",     0, "Debit minus Credit", "Project Plan Task", "Ledger Account", "Ledger Hierarchy Level 5", "Payroll", "Opex", None),
    ("OPEX.xlsx",       "Data#5",     0, "Debit minus Credit", "Project Plan Task", "Ledger Account", "Ledger Hierarchy Level 5", "Payroll", "Opex", None),
    ("OPEX.xlsx",       "Data#6",     0, "Amount",             "Project",           "Ledger Account", "Ledger Account Summary (Level 5)", "WD4_COGS", "Opex", None),
    ("CAPEX.xlsx",      "JL details", 0, "Debit minus Credit", "Project Code",      "Ledger Account", "Ledger Hierarchy Level 5", "CAPEX",   "Capex", None),
]
GAMING_TABS = ["項目1", "項目2", "項目3", "項目5", "項目6", "項目7", "項目8", "項目9", "項目10"]


def leadsheet_ng_map():
    fp = find_file("OPEX.xlsx")
    df = pd.read_excel(fp, sheet_name="Leadsheet", header=0, dtype=object)
    c_no, c_nat = col(df, "承批公司項目序號"), col(df, "項目性質")
    m = {}
    if c_no and c_nat:
        for _, r in df.iterrows():
            k = seq(r[c_no])
            if k and str(r[c_nat]).strip() not in ("", "nan"):
                m.setdefault(k, str(r[c_nat]).strip())
    return m


def main():
    ngmap = leadsheet_ng_map()
    print(f"Leadsheet NG map: {len(ngmap)} 項目")
    out = []
    for fname, sheet, hdr, amt, proj, ac, hier, src, capex, ngc in SOURCES:
        fp = find_file(fname)
        if not fp: print(f"[{src}/{sheet}] {fname} missing"); continue
        try:
            df = pd.read_excel(fp, sheet_name=sheet, header=hdr, dtype=object)
        except Exception as e:
            print(f"[{src}/{sheet}] read failed: {e}"); continue
        c_amt = col(df, amt)
        if not c_amt: print(f"[{src}/{sheet}] amount {amt!r} missing"); continue
        c_pr, c_ac, c_hi = col(df, proj), col(df, ac), col(df, hier)
        pj = (df[c_pr].map(seq) if (c_pr and proj in ("承批公司項目序號",)) else
              (df[c_pr].astype(str).str.strip() if c_pr else ""))
        ng = pj.map(lambda k: ngmap.get(k, "")) if hasattr(pj, "map") else ""
        sub = pd.DataFrame({
            "Debit minus Credit": num(df[c_amt]).values,
            "Source": src,
            "Project_code": pj.values if hasattr(pj, "values") else pj,
            "Section.1": ng.values if hasattr(ng, "values") else "",
            "Ledger Account": df[c_ac].astype(str).str.strip().values if c_ac else "",
            "Ledger Hierarchy Level 5": df[c_hi].astype(str).str.strip().values if c_hi else "",
        })
        out.append(sub)
        print(f"[{src}/{sheet}] {len(sub):,} rows  Σ={sub['Debit minus Credit'].sum():,.0f}")
    # gaming capex tabs (tab name = 項目序號, NG=博彩)
    fpg = find_file("MGM-gaming.xlsx")
    if fpg:
        for tab in GAMING_TABS:
            try:
                df = pd.read_excel(fpg, sheet_name=tab, header=0, dtype=object)
            except Exception:
                continue
            c_amt = col(df, "Debit minus Credit")
            if not c_amt: continue
            c_ac = col(df, "Ledger Account"); c_hi = col(df, "Ledger Hierarchy Level 5")
            sub = pd.DataFrame({
                "Debit minus Credit": num(df[c_amt]).values,
                "Source": "CAPEX",
                "Project_code": "項目" + str(int(re.sub(r"\D", "", tab)) + 200),  # 博彩 序號 +200 → match golden
                "Section.1": "博彩娛樂場場地的優化",
                "Ledger Account": df[c_ac].astype(str).str.strip().values if c_ac else "",
                "Ledger Hierarchy Level 5": df[c_hi].astype(str).str.strip().values if c_hi else "",
            })
            out.append(sub)
            print(f"[CAPEX/{tab}] {len(sub):,} rows  Σ={sub['Debit minus Credit'].sum():,.0f}")
    if not out: print("nothing built"); return
    allr = pd.concat(out, ignore_index=True)
    res = ROOT / "data" / "mgm" / "raw"; res.mkdir(parents=True, exist_ok=True)
    fpo = res / "mgm_23_raw.xlsx"
    with pd.ExcelWriter(fpo, engine="openpyxl") as w:
        allr.to_excel(w, sheet_name="combine", index=False)
    print(f"\n✓ {len(allr):,} rows → {fpo.relative_to(ROOT)}  Σ={allr['Debit minus Credit'].sum():,.0f}")
    print(allr.groupby("Source")["Debit minus Credit"].sum().round(0).to_string())
    print("NG blank rows:", int((allr["Section.1"].isin(["", "nan"]) | allr["Section.1"].isna()).sum()))
    print("\n(raw 係做 V/H 比例;總額由 golden 866M tie,唔需要 raw Σ = 866M)")


if __name__ == "__main__":
    main()
