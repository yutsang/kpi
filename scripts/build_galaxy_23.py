"""Prebuild Galaxy 2023 raw — union the 9 取數 sources into ONE sheet matching galaxy's
pipeline schema (Reported Amount(MOP) | Capex/Opex | Project | NG11 Category | Account Code
| Account Description | Description | Source | Submit No.). Total ties Cover Σ = 3.384B.

Then conf/company_1 yearly_sources 2023 points at this file; the pipeline classifies V (step2
within NG) + H (predominant_rules on Account Code + Source). NG = 'NG11 Category' (given).

Run (Windows):
  python scripts/build_galaxy_23.py --pw dicj_kpmg
Output → data/galaxy/raw/2023/galaxy_23_raw.xlsx  (sheet 'combine')
"""
from __future__ import annotations
import argparse, io, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent


def find_file(name):
    p = Path(name)
    if p.exists(): return p
    d = ROOT / "data"
    hits = list(d.glob(f"*/raw/{name}")) + list(d.rglob(name)) + list(d.rglob(f"*{name}*"))
    exact = [h for h in hits if h.name == name]
    return (exact or hits)[0] if (exact or hits) else None


def handle(fp, pw):
    if not pw: return fp
    import msoffcrypto
    buf = io.BytesIO()
    with open(fp, "rb") as f:
        of = msoffcrypto.OfficeFile(f); of.load_key(password=pw); of.decrypt(buf)
    buf.seek(0); return buf


def rd(fp, sheet, header, pw):
    return pd.read_excel(handle(fp, pw), sheet_name=sheet, header=header, dtype=object)


def col(df, name):
    if name is None: return None
    if name in df.columns: return name
    for c in df.columns:
        if str(c).strip() == str(name).strip(): return c
    return None


def num(s): return pd.to_numeric(s, errors="coerce")


# per-source: (sheet, header, amt, capex flag, ac, ad, proj, ng, desc, filter, file=sap|book)
def SOURCES(pw_set):
    S = [
        dict(src="SAP_opex",     fk="sap", sheet="SAP Record-non gaming", hdr=0, capex="Opex",
             amt="Reported Amount (MOP)", ac="Account Code", ad="Account Description",
             proj="Project", ng="NG11 Category", desc="Description"),
        dict(src="SAP_capex_ng", fk="sap", sheet="SAP Record-non gaming", hdr=0, capex="Capx",
             amt="Reported Amount (MOP).1", ac="Account Code.1", ad="Account Description.1",
             proj="Project", ng="NG11 Category", desc="Description.1"),
        dict(src="PMS",          fk="sap", sheet="PMS", hdr=0, capex="Opex",
             amt="Amount (MOP)", ac="Account Code", ad="Account Description",
             proj="Project", ng="NG11 Category", desc="Description"),
        dict(src="VMS",          fk="sap", sheet="VMS", hdr=3, capex="Opex",
             amt="Amount (Reported Amt) (MOP)", ac=None, ad="Descriptions",
             proj="Project", ng="NG11 Category", desc="Voucher Discroption"),
        dict(src="VR",           fk="sap", sheet="VR Record", hdr=3, capex="Opex",
             amt="Amount", ac=None, ad="Description.1", proj="Project", ng="NG11 Category", desc="Description"),
        dict(src="Simulation",   fk="sap", sheet="Simulation&Pre-Opening", hdr=2, capex="Opex",
             amt="Reported Amount", ac="Account Code", ad="Account Description",
             proj=None, ng=None, desc="Event Descriptions", filt=("Nature", "Simulation"), cat="項目類別"),
        dict(src="PreOpening",   fk="sap", sheet="Simulation&Pre-Opening", hdr=2, capex="Opex",
             amt="Reported Amount", ac="Account Code", ad="Account Description",
             proj=None, ng=None, desc="Description", filt=("Nature", "Pre-Opening")),
        dict(src="Corporate",    fk="sap", sheet="Corporate support", hdr=18, capex="Opex",
             amt_letter="R", ac=None, ad=None, proj=None, ng=None, desc=None),
        dict(src="Other",        fk="sap", sheet="Other Source Detail", hdr=0, capex="Opex",
             amt="Investment Amount", ac="GL Account", ad="GL Account Description",
             proj="Project", ng="NG11 Category", desc="Description"),
    ]
    if pw_set:
        S.append(dict(src="SAP_capex_gam", fk="book", sheet="SAP Gaming", hdr=1, capex="Capx",
                      amt="Reported Amount (MOP).1", ac="Account Code.1", ad="Account Description.1",
                      proj="Project", ng="NG11 Category", desc="Line Item Description"))
    return S

SUBPROJ = r"[A-Za-z]\d+\.\d{1,}"


def letter_to_idx(letter):
    idx = 0
    for ch in str(letter).strip().upper():
        if "A" <= ch <= "Z": idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def build(files, pw):
    out = []
    for s in SOURCES(bool(pw)):
        fp = files[s["fk"]]
        if not fp: print(f"[{s['src']}] file missing"); continue
        try:
            df = rd(fp, s["sheet"], s["hdr"], pw if s["fk"] == "book" else None)
        except Exception as e:
            print(f"[{s['src']}] read failed: {e}"); continue
        if s.get("filt"):
            fc = col(df, s["filt"][0])
            if fc: df = df[df[fc].astype(str).str.strip().str.lower() == s["filt"][1].lower()]
        # amount
        amt = col(df, s.get("amt"))
        if amt: a = num(df[amt])
        elif s.get("amt_letter"):
            li = letter_to_idx(s["amt_letter"]); a = num(df.iloc[:, li]) if li < len(df.columns) else None
        else: a = None
        if a is None: print(f"[{s['src']}] no amount — skip"); continue
        n = len(df)
        ac, ad = col(df, s.get("ac")), col(df, s.get("ad"))
        proj, ng, desc = col(df, s.get("proj")), col(df, s.get("ng")), col(df, s.get("desc"))
        cat = col(df, s.get("cat"))
        sub = pd.DataFrame({
            "Reported Amount(MOP)": a.values,
            "Capex/Opex": s["capex"],
            "Project": (df[proj].astype(str).str.replace(r"[\r\n]+", " ", regex=True).str.strip().values
                        if proj else (f"{s['src']}",) * n),
            "NG11 Category": (df[ng].astype(str).str.strip().values if ng
                              else (df[cat].astype(str).str.strip().values if cat else ("",) * n)),
            "Account Code": df[ac].astype(str).str.strip().values if ac else "",
            "Account Description": df[ad].astype(str).str.strip().values if ad else f"{s['src']}",
            "Description": df[desc].astype(str).str.strip().values if desc else f"{s['src']}",
            "Source": s["src"],
        })
        # drop grand-total / blank rows (Corporate has a 277.9M total row + DICJ rows)
        if s["src"] == "Corporate":
            spcol = next((c for c in df.columns if df[c].dropna().astype(str).str.fullmatch(SUBPROJ).mean() > 0.5), None)
            if spcol is not None:
                keep = df[spcol].astype(str).str.fullmatch(SUBPROJ).fillna(False).values
                sub = sub[keep]; sub["Project"] = df.loc[keep, spcol].astype(str).str.strip().values
        out.append(sub)
        print(f"[{s['src']}] {len(sub):,} rows  Σ={sub['Reported Amount(MOP)'].sum():,.0f}")
    # ── 沒有SAP Record residual: Cover Amount(I) − Σ(9 components J..R) per Cover row.
    #    Carries Cover 'Descriptions' (F&B/Staff Costs/Hotel/Utilities/Corporate Support…)
    #    as Account Description so H rules classify it. Ties total to Cover 3.384B. ──
    try:
        cv = rd(files["sap"], "Cover", 2, None)
        covI = num(cv.iloc[:, letter_to_idx("I")]).fillna(0)
        comp9 = sum(num(cv.iloc[:, letter_to_idx(l)]).fillna(0) for l in "JKLMNOPQR")
        resid = covI - comp9
        keep = resid.abs() > 1
        ng, proj, desc = col(cv, "NG11 Category"), col(cv, "Project"), col(cv, "Descriptions")
        nosap = pd.DataFrame({
            "Reported Amount(MOP)": resid[keep].values,
            "Capex/Opex": "Opex",
            "Project": cv.loc[keep, proj].astype(str).str.replace(r"[\r\n]+", " ", regex=True).str.strip().values,
            "NG11 Category": cv.loc[keep, ng].astype(str).str.strip().values,
            "Account Code": "",
            "Account Description": cv.loc[keep, desc].astype(str).str.strip().values,  # cost nature → H
            "Description": cv.loc[keep, desc].astype(str).str.strip().values,
            "Source": "NoSAPRecord",
        })
        out.append(nosap)
        print(f"[NoSAPRecord] {len(nosap):,} rows  Σ={nosap['Reported Amount(MOP)'].sum():,.0f}")
    except Exception as e:
        print(f"[NoSAPRecord] residual pass failed: {e}")

    if not out: print("nothing built"); return
    allr = pd.concat(out, ignore_index=True)
    allr["Submit No."] = ["G23_" + str(i) for i in range(len(allr))]   # unique_id, no project name
    res = ROOT / "data" / "galaxy" / "raw"   # step0 resolves yearly_sources from raw/ (NOT raw/2023/)
    res.mkdir(parents=True, exist_ok=True)
    fp_out = res / "galaxy_23_raw.xlsx"
    with pd.ExcelWriter(fp_out, engine="openpyxl") as w:
        allr.to_excel(w, sheet_name="combine", index=False)
    print(f"\n✓ {len(allr):,} rows → {fp_out.relative_to(ROOT)}  Σ={allr['Reported Amount(MOP)'].sum():,.0f}  (expect Cover 3,384,404,144)")
    print("  by Source:"); print(allr.groupby("Source")["Reported Amount(MOP)"].sum().round(0).to_string())
    print("  NG blank rows:", int((allr["NG11 Category"].isin(["", "nan", "None"]) | allr["NG11 Category"].isna()).sum()))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sap", default="galaxy_sap.xlsx")
    ap.add_argument("--book", default="galaxy_raw_databook.xlsx")
    ap.add_argument("--pw", default=None, help="password for the encrypted databook (gaming capex)")
    a = ap.parse_args()
    files = {"sap": find_file(a.sap), "book": find_file(a.book)}
    print(f"sap={files['sap']}  book={files['book']}  pw={'set' if a.pw else 'none'}")
    build(files, a.pw)


if __name__ == "__main__":
    main()
