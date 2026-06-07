"""Prebuild MGM 2024 raw — 2024 has row-level JE with NG inline (WD1 'Section'), so it's
BOTTOM-UP (no golden, unlike 2023). Union the granular sheets into the 2025 'combine' schema.

Source = 2024年JE.xlsx (data/mgm/raw/2024/). Use the GRANULAR sheets (NOT the consolidated
'2024 Journal line' which overlaps WD1+Capex → double-count):
  WD1 (opex, Section=NG inline) · Capex · PM (Patron comp) · WD2-Allocation
NG (Section.1) = 'Section' with the ' - 活動/設施' suffix stripped.
(Payroll = Retendering allocation file — added later if needed; flagged as missing.)

Run (Windows):
  python scripts/build_mgm_24_raw.py
Output → data/mgm/raw/mgm_24_raw.xlsx  (sheet 'combine')
"""
from __future__ import annotations
import re, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent


def find_file(name):
    for p in [ROOT / "data" / "mgm" / "raw" / "2024" / name, ROOT / "data" / "mgm" / "raw" / name]:
        if p.exists(): return p
    hits = list((ROOT / "data").rglob(name))
    return hits[0] if hits else None


def col(df, name):
    if name is None: return None
    if name in df.columns: return name
    for c in df.columns:
        if str(c).strip() == str(name).strip(): return c
    for c in df.columns:
        if str(name).strip() in str(c).strip():
            return c
    return None


def num(s): return pd.to_numeric(s, errors="coerce").fillna(0.0)


def strip_ng(v):
    s = str(v).strip()
    s = re.sub(r"\s*[-–—]\s*(活動|設施).*$", "", s)   # '吸引外國客源 - 活動' → '吸引外國客源'
    return "" if s in ("nan", "None") else s


# (sheet, header, amount, project, account, hier, ng, source)
SOURCES = [
    ("WD1",            0, "Debit minus Credit", "Project Plan Task", "Ledger Account", "Ledger Hierarchy Level 5", "Section", "WD1"),
    ("WD2-Allocation", 0, "Debit minus Credit", "Project Plan Task", "Ledger Account", "Ledger Hierarchy Level 5", "Section", "WD2"),
    ("Capex",          0, "Debit minus Credit", "Project Code",      "Ledger Account", "Ledger Hierarchy Level 5", "Section", "CAPEX"),
    ("PM",             0, "Amount",             None,                "Item Type",      "Promotion Type",           "Section", "PM"),
]


def main():
    fp = find_file("2024年JE.xlsx")
    if not fp:
        print("X 2024年JE.xlsx not found under data/mgm/raw/2024/"); return
    print(f"file: {fp}")
    out = []
    for sheet, hdr, amt, proj, ac, hier, ng, src in SOURCES:
        try:
            df = pd.read_excel(fp, sheet_name=sheet, header=hdr, dtype=object)
        except Exception as e:
            print(f"[{src}/{sheet}] read failed: {e}"); continue
        c_amt = col(df, amt)
        if not c_amt: print(f"[{src}/{sheet}] amount {amt!r} missing — cols: {list(df.columns)[:12]}"); continue
        c_pr, c_ac, c_hi, c_ng = col(df, proj), col(df, ac), col(df, hier), col(df, ng)
        sub = pd.DataFrame({
            "Debit minus Credit": num(df[c_amt]).values,
            "Source": src,
            "Project_code": df[c_pr].astype(str).str.strip().values if c_pr else "",
            "Section.1": (df[c_ng].map(strip_ng).values if c_ng else ""),
            "Ledger Account": df[c_ac].astype(str).str.strip().values if c_ac else "",
            "Ledger Hierarchy Level 5": df[c_hi].astype(str).str.strip().values if c_hi else "",
        })
        out.append(sub)
        print(f"[{src}/{sheet}] {len(sub):,} rows  Σ={sub['Debit minus Credit'].sum():,.0f}")
    if not out: print("nothing built"); return
    allr = pd.concat(out, ignore_index=True)
    # ── NG (Section.1) backfill — only WD1 carries inline 'Section'; Capex/PM/WD2 tabs lack it
    #    → blank NG → step2 candidate collapse → V_OTHER (38% of 2024). Two fills:
    #    (1) PM = Patron comp → 吸引外國客源 (NG1; matches the 2025 WD5_Patron rule).
    #    (2) propagate each Project_code's modal non-blank NG to its blank siblings — recovers
    #        Capex/WD2 rows that share WD1's 項目 token; harmless no-op where tokens differ.
    def _isblank(s): return s.astype(str).str.strip().isin(["", "nan", "None"])
    allr.loc[_isblank(allr["Section.1"]) & allr["Source"].eq("PM"), "Section.1"] = "吸引外國客源"
    _known = allr[~_isblank(allr["Section.1"])]
    if len(_known) and "Project_code" in allr.columns:
        _pmap = _known.groupby("Project_code")["Section.1"].agg(lambda s: s.value_counts().index[0])
        _bl = _isblank(allr["Section.1"])
        _fill = allr.loc[_bl, "Project_code"].map(_pmap)
        allr.loc[_bl, "Section.1"] = _fill.where(_fill.notna(), allr.loc[_bl, "Section.1"])
    print(f"  [NG backfill] PM→NG1 + project-modal fill; remaining NG-blank rows: "
          f"{int(_isblank(allr['Section.1']).sum()):,}")
    res = ROOT / "data" / "mgm" / "raw"; res.mkdir(parents=True, exist_ok=True)
    fpo = res / "mgm_24_raw.xlsx"
    with pd.ExcelWriter(fpo, engine="openpyxl") as w:
        allr.to_excel(w, sheet_name="combine", index=False)
    print(f"\n✓ {len(allr):,} rows → {fpo.relative_to(ROOT)}  Σ={allr['Debit minus Credit'].sum():,.0f}")
    print(allr.groupby("Source")["Debit minus Credit"].sum().round(0).to_string())
    print("NG blank rows:", int((allr["Section.1"].isin(["", "nan"]) | allr["Section.1"].isna()).sum()))
    print("\n⚠️ v1 bottom-up — 無 payroll(喺 Retendering file);若總額同預期差太遠,睇下要唔要 2024 Journal line 代替 WD1+Capex(免/補 double-count)。")


if __name__ == "__main__":
    main()
